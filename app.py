from __future__ import annotations

import io
import json
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from functools import wraps
from pathlib import Path
from typing import Any

import firebase_admin
from firebase_admin import credentials, db
from dotenv import load_dotenv
from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    send_file,
    url_for,
)
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from openpyxl import Workbook
from werkzeug.security import check_password_hash, generate_password_hash


BASE_DIR = Path(__file__).resolve().parent

load_dotenv(BASE_DIR / ".env")


app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-me-in-production")
app.config["REMEMBER_COOKIE_DURATION"] = timedelta(days=int(os.environ.get("REMEMBER_DAYS", "30")))
login_manager = LoginManager(app)
login_manager.login_view = "login"


CATEGORIES = ["Food", "Travel", "Fuel", "Shopping", "Office", "Utilities", "Other"]


# ---------------------------------------------------------------------------
# Firebase database connection
# ---------------------------------------------------------------------------
def firebase_database_url() -> str:
    value = os.environ.get(
        "FIREBASE_DATABASE_URL",
        "https://personal-f48e3-default-rtdb.firebaseio.com/",
    ).strip()
    return value or "https://personal-f48e3-default-rtdb.firebaseio.com/"


def firebase_credentials_source() -> str | dict[str, Any]:
    raw = os.environ.get("FIREBASE_CREDENTIALS_JSON", "").strip()
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RuntimeError("FIREBASE_CREDENTIALS_JSON is not valid JSON.") from exc

    path_value = os.environ.get("FIREBASE_CREDENTIALS_PATH", "").strip()
    path = Path(path_value) if path_value else BASE_DIR / "firebase-credentials.json"
    if not path.exists():
        raise RuntimeError(
            "Firebase credentials are not configured. Set FIREBASE_CREDENTIALS_JSON "
            "or FIREBASE_CREDENTIALS_PATH to a service-account JSON file."
        )
    return str(path)


def initialize_firebase() -> None:
    try:
        firebase_admin.get_app()
        return
    except ValueError:
        pass

    cred_source = firebase_credentials_source()
    options = {"databaseURL": firebase_database_url()}
    firebase_admin.initialize_app(credentials.Certificate(cred_source), options)


def firebase_ref(path: str = ""):
    initialize_firebase()
    return db.reference(path)


def utc_now() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def serialize_datetime(value: datetime | None) -> str | None:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def parse_datetime_value(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value if value.tzinfo is not None else value.replace(tzinfo=timezone.utc)
    if isinstance(value, str):
        normalized = value.replace("Z", "+00:00")
        try:
            parsed = datetime.fromisoformat(normalized)
        except ValueError:
            return None
        return parsed if parsed.tzinfo is not None else parsed.replace(tzinfo=timezone.utc)
    return None


def normalize_date(value: Any) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, str):
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def coerce_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return parse_bool(str(value))


def table_rows(table: str) -> list[dict[str, Any]]:
    raw = firebase_ref(table).get() or {}
    if isinstance(raw, list):
        rows = [row for row in raw if isinstance(row, dict)]
    elif isinstance(raw, dict):
        rows = []
        for key, value in raw.items():
            if not isinstance(value, dict):
                continue
            row = dict(value)
            row.setdefault("id", parse_int(str(key), 0) or 0)
            rows.append(row)
    else:
        rows = []
    rows.sort(key=lambda row: row.get("id", 0))
    return rows


def get_record(table: str, record_id: int) -> dict[str, Any] | None:
    row = firebase_ref(f"{table}/{record_id}").get()
    if isinstance(row, dict):
        row.setdefault("id", record_id)
        return row
    return None


def set_record(table: str, record_id: int, data: dict[str, Any]) -> None:
    firebase_ref(f"{table}/{record_id}").set(data)


def delete_record(table: str, record_id: int) -> None:
    firebase_ref(f"{table}/{record_id}").delete()


def next_record_id(table: str) -> int:
    counter_ref = firebase_ref(f"meta/counters/{table}")

    def increment(current):
        return int(current or 0) + 1

    return int(counter_ref.transaction(increment))


def init_db() -> None:
    initialize_firebase()
    counters_ref = firebase_ref("meta/counters")
    counters = counters_ref.get() or {}
    if not isinstance(counters, dict):
        counters = {}
    changed = False
    for name in ("users", "expenses", "audit"):
        if counters.get(name) is None:
            counters[name] = 0
            changed = True
    if changed:
        counters_ref.set(counters)


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
@dataclass
class AppUser(UserMixin):
    id: int
    username: str
    password_hash: str
    is_admin: bool
    is_active_account: bool
    created_at: datetime | None

    @property
    def is_active(self) -> bool:
        return self.is_active_account

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


@dataclass
class ExpenseRecord:
    id: int
    user_id: int
    amount: int
    description: str
    category: str
    expense_date: date
    created_at: datetime | None
    updated_at: datetime | None
    username: str = ""


@dataclass
class AuditEntry:
    id: int
    actor_id: int | None
    action: str
    details: str
    created_at: datetime | None
    actor_name: str = ""


def row_to_user(row) -> AppUser:
    return AppUser(
        id=parse_int(row["id"], 0) or 0,
        username=row["username"],
        password_hash=row["password_hash"],
        is_admin=coerce_bool(row.get("is_admin")),
        is_active_account=coerce_bool(row.get("is_active_account")),
        created_at=parse_datetime_value(row.get("created_at")),
    )


def row_to_expense(row) -> ExpenseRecord:
    return ExpenseRecord(
        id=parse_int(row["id"], 0) or 0,
        user_id=parse_int(row["user_id"], 0) or 0,
        amount=parse_int(row["amount"], 0) or 0,
        description=row["description"],
        category=row["category"],
        expense_date=normalize_date(row.get("expense_date")) or date.today(),
        created_at=parse_datetime_value(row.get("created_at")),
        updated_at=parse_datetime_value(row.get("updated_at")),
        username=row.get("username", ""),
    )


def row_to_audit(row) -> AuditEntry:
    return AuditEntry(
        id=parse_int(row["id"], 0) or 0,
        actor_id=parse_int(row["actor_id"]) if row.get("actor_id") is not None else None,
        action=row["action"],
        details=row["details"],
        created_at=parse_datetime_value(row.get("created_at")),
        actor_name=row.get("actor_name", "System"),
    )


# ---------------------------------------------------------------------------
# Parsing / formatting helpers
# ---------------------------------------------------------------------------
def parse_bool(value: str | None) -> bool:
    return value in {"1", "true", "True", "on", "yes"}


def parse_int(value: str | None, default: int | None = None) -> int | None:
    if value in (None, ""):
        return default
    try:
        return int(value)
    except ValueError:
        return default


def parse_date(value: str | None, default: date | None = None) -> date | None:
    if value in (None, ""):
        return default
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return default


def money(value: int | None) -> str:
    if value is None:
        return "Rs. 0"
    return f"Rs. {value:,}"


# ---------------------------------------------------------------------------
# Serialization helpers
# ---------------------------------------------------------------------------
def user_to_dict(user: AppUser) -> dict[str, Any]:
    return {
        "id": user.id,
        "username": user.username,
        "password_hash": user.password_hash,
        "is_admin": user.is_admin,
        "is_active_account": user.is_active_account,
        "created_at": serialize_datetime(user.created_at) or serialize_datetime(utc_now()),
    }


def expense_to_dict(expense: ExpenseRecord) -> dict[str, Any]:
    return {
        "id": expense.id,
        "user_id": expense.user_id,
        "amount": expense.amount,
        "description": expense.description,
        "category": expense.category,
        "expense_date": expense.expense_date.isoformat(),
        "created_at": serialize_datetime(expense.created_at) or serialize_datetime(utc_now()),
        "updated_at": serialize_datetime(expense.updated_at) or serialize_datetime(utc_now()),
    }


def audit_to_dict(entry: AuditEntry) -> dict[str, Any]:
    return {
        "id": entry.id,
        "actor_id": entry.actor_id,
        "action": entry.action,
        "details": entry.details,
        "created_at": serialize_datetime(entry.created_at) or serialize_datetime(utc_now()),
    }


# ---------------------------------------------------------------------------
# User data access
# ---------------------------------------------------------------------------
def get_users() -> list[AppUser]:
    return [row_to_user(row) for row in table_rows("users")]


def get_user_by_id(user_id: int) -> AppUser | None:
    row = get_record("users", user_id)
    return row_to_user(row) if row else None


def get_user_by_username(username: str, *, is_admin: bool | None = None) -> AppUser | None:
    for user in get_users():
        if user.username != username:
            continue
        if is_admin is not None and user.is_admin != is_admin:
            continue
        return user
    return None


def active_admin_count() -> int:
    return sum(1 for user in get_users() if user.is_admin and user.is_active_account)


def insert_user(
    username: str,
    password_hash: str,
    is_admin: bool,
    is_active_account: bool = True,
) -> AppUser:
    user = AppUser(
        id=next_record_id("users"),
        username=username,
        password_hash=password_hash,
        is_admin=is_admin,
        is_active_account=is_active_account,
        created_at=utc_now(),
    )
    set_record("users", user.id, user_to_dict(user))
    return user


def save_user(user: AppUser) -> None:
    """Persist changes to an existing user (matched by id)."""
    if user.created_at is None:
        user.created_at = utc_now()
    set_record("users", user.id, user_to_dict(user))


# ---------------------------------------------------------------------------
# Expense data access
# ---------------------------------------------------------------------------
def get_expenses() -> list[ExpenseRecord]:
    users = {user.id: user.username for user in get_users()}
    expenses = []
    for row in table_rows("expenses"):
        record = row_to_expense(row)
        record.username = users.get(record.user_id, "Unknown")
        expenses.append(record)
    return expenses


def get_expense_by_id(expense_id: int) -> ExpenseRecord | None:
    row = get_record("expenses", expense_id)
    if row is None:
        return None
    expense = row_to_expense(row)
    user = get_user_by_id(expense.user_id)
    expense.username = user.username if user else "Unknown"
    return expense


def insert_expense(
    user_id: int,
    amount: int,
    description: str,
    category: str,
    expense_date: date,
) -> ExpenseRecord:
    user = get_user_by_id(user_id)
    expense = ExpenseRecord(
        id=next_record_id("expenses"),
        user_id=user_id,
        amount=amount,
        description=description,
        category=category,
        expense_date=expense_date,
        created_at=utc_now(),
        updated_at=utc_now(),
        username=user.username if user else "Unknown",
    )
    set_record("expenses", expense.id, expense_to_dict(expense))
    return expense


def save_expense_record(expense: ExpenseRecord) -> None:
    """Persist changes to an existing expense (matched by id)."""
    if expense.created_at is None:
        expense.created_at = utc_now()
    expense.updated_at = utc_now()
    set_record("expenses", expense.id, expense_to_dict(expense))


def delete_expense_record(expense_id: int) -> ExpenseRecord | None:
    row = get_record("expenses", expense_id)
    if row is None:
        return None
    delete_record("expenses", expense_id)
    expense = row_to_expense(row)
    user = get_user_by_id(expense.user_id)
    expense.username = user.username if user else "Unknown"
    return expense


# ---------------------------------------------------------------------------
# Audit data access
# ---------------------------------------------------------------------------
def get_audit_entries(limit: int = 200) -> list[AuditEntry]:
    users = {user.id: user.username for user in get_users()}
    entries = []
    for row in sorted(table_rows("audit"), key=lambda item: item.get("id", 0), reverse=True)[:limit]:
        entry = row_to_audit(row)
        entry.actor_name = users.get(entry.actor_id, "System") if entry.actor_id is not None else "System"
        entries.append(entry)
    return entries


def log_action(action: str, details: str, actor_id: int | None = None) -> None:
    entry = AuditEntry(
        id=next_record_id("audit"),
        actor_id=actor_id,
        action=action,
        details=details[:500],
        created_at=utc_now(),
        actor_name="System",
    )
    set_record("audit", entry.id, audit_to_dict(entry))


# ---------------------------------------------------------------------------
# Admin bootstrap
# ---------------------------------------------------------------------------
def ensure_admin_account() -> None:
    """Keep the admin account in sync with the configured environment variables.

    ``ADMIN_USERNAME`` and ``ADMIN_PASSWORD`` are treated as the source of truth
    for administrator access in both local and production (Render) environments.
    On startup the admin account is created if it is missing, reactivated if it
    was disabled, and its password is reset whenever the configured value
    changes.
    """
    username = os.environ.get("ADMIN_USERNAME", "admin").strip() or "admin"
    password = os.environ.get("ADMIN_PASSWORD", "admin12345")

    admin = get_user_by_username(username, is_admin=True)
    if admin is None:
        insert_user(username, generate_password_hash(password), is_admin=True, is_active_account=True)
        return

    changed = False
    if not admin.is_active_account:
        admin.is_active_account = True
        changed = True
    if not admin.check_password(password):
        admin.set_password(password)
        changed = True
    if changed:
        save_user(admin)


# ---------------------------------------------------------------------------
# Access control / query helpers
# ---------------------------------------------------------------------------
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin:
            abort(403)
        return fn(*args, **kwargs)

    return wrapper


def filter_expenses(expenses: list[ExpenseRecord]) -> list[ExpenseRecord]:
    start_date = parse_date(request.args.get("start_date"))
    end_date = parse_date(request.args.get("end_date"))
    category = request.args.get("category", "").strip()
    query = request.args.get("q", "").strip().lower()
    min_amount = parse_int(request.args.get("min_amount"))
    max_amount = parse_int(request.args.get("max_amount"))
    user_id = parse_int(request.args.get("user_id")) if current_user.is_admin else None

    filtered = []
    for expense in expenses:
        if start_date and expense.expense_date < start_date:
            continue
        if end_date and expense.expense_date > end_date:
            continue
        if category and expense.category != category:
            continue
        if query:
            haystack = " ".join(
                [
                    expense.description.lower(),
                    expense.category.lower(),
                    expense.username.lower(),
                ]
            )
            if query not in haystack:
                continue
        if min_amount is not None and expense.amount < min_amount:
            continue
        if max_amount is not None and expense.amount > max_amount:
            continue
        if user_id is not None and expense.user_id != user_id:
            continue
        filtered.append(expense)
    return filtered


def current_expenses() -> list[ExpenseRecord]:
    expenses = get_expenses()
    if current_user.is_admin:
        return expenses
    return [expense for expense in expenses if expense.user_id == current_user.id]


def dashboard_totals(expenses: list[ExpenseRecord]) -> dict[str, int | str]:
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    month_start = today.replace(day=1)
    return {
        "today": sum(expense.amount for expense in expenses if expense.expense_date == today),
        "week": sum(expense.amount for expense in expenses if expense.expense_date >= week_start),
        "month": sum(expense.amount for expense in expenses if expense.expense_date >= month_start),
        "total": sum(expense.amount for expense in expenses),
        "today_label": today.strftime("%d %b %Y"),
    }


def monthly_breakdown(expenses: list[ExpenseRecord]) -> dict[str, int]:
    totals = {category: 0 for category in CATEGORIES}
    for expense in expenses:
        totals[expense.category] = totals.get(expense.category, 0) + expense.amount
    return totals


def recent_days_series(expenses: list[ExpenseRecord], days: int = 14):
    end = date.today()
    start = end - timedelta(days=days - 1)
    day_map: dict[str, int] = {}
    current = start
    while current <= end:
        day_map[current.isoformat()] = 0
        current += timedelta(days=1)
    for expense in expenses:
        if start <= expense.expense_date <= end:
            key = expense.expense_date.isoformat()
            day_map[key] = day_map.get(key, 0) + expense.amount
    labels = [datetime.strptime(day, "%Y-%m-%d").strftime("%d %b") for day in day_map.keys()]
    values = list(day_map.values())
    return labels, values


def chart_payload(expenses: list[ExpenseRecord]):
    category_totals = monthly_breakdown(expenses)
    trend_labels, trend_values = recent_days_series(expenses)
    return {
        "category_labels": list(category_totals.keys()),
        "category_values": list(category_totals.values()),
        "trend_labels": trend_labels,
        "trend_values": trend_values,
    }


def manual_pagination(items: list[ExpenseRecord], page: int, per_page: int = 20):
    total = len(items)
    pages = max(1, (total + per_page - 1) // per_page)
    page = max(1, min(page, pages))
    start = (page - 1) * per_page
    end = start + per_page

    class Pagination:
        def __init__(self):
            self.items = items[start:end]
            self.page = page
            self.pages = pages
            self.has_prev = page > 1
            self.has_next = page < pages
            self.prev_num = page - 1
            self.next_num = page + 1

    return Pagination()


def build_url(endpoint: str, **updates):
    args = request.args.to_dict(flat=True)
    args.pop("page", None)
    args.update({key: value for key, value in updates.items() if value is not None and value != ""})
    return url_for(endpoint, **args)


@login_manager.user_loader
def load_user(user_id: str):
    user = get_user_by_id(int(user_id))
    return user if user and user.is_active_account else None


@app.template_filter("money")
def money_filter(value):
    return money(value)


@app.context_processor
def inject_globals():
    return {
        "categories": CATEGORIES,
        "money": money,
        "today": date.today(),
        "app_name": "Expense Tracker",
        "build_url": build_url,
    }


@app.before_request
def make_session_permanent():
    if current_user.is_authenticated:
        session.permanent = True


@app.route("/")
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    users = get_users()
    admin_user = get_user_by_username(os.environ.get("ADMIN_USERNAME", "admin"), is_admin=True)
    return render_template(
        "home.html",
        admin_user=admin_user,
        user_count=len(users),
        expense_count=len(get_expenses()),
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = get_user_by_username(username, is_admin=False)
        if user and user.is_active_account and user.check_password(password):
            login_user(user, remember=True)
            log_action("login", f"User {user.username} logged in", user.id)
            return redirect(url_for("dashboard"))
        flash("The username or password you entered is incorrect. Please try again.", "error")
    return render_template("login.html", mode="user")


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if current_user.is_authenticated and current_user.is_admin:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = get_user_by_username(username, is_admin=True)
        if user and user.is_active_account and user.check_password(password):
            login_user(user, remember=True)
            log_action("admin_login", f"Admin {user.username} logged in", user.id)
            return redirect(url_for("dashboard"))
        flash("The administrator username or password you entered is incorrect. Please try again.", "error")
    return render_template("login.html", mode="admin")


@app.route("/logout")
@login_required
def logout():
    log_action("logout", f"{current_user.username} logged out", current_user.id)
    logout_user()
    flash("You have been signed out successfully.", "success")
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    items = filter_expenses(current_expenses())
    totals = dashboard_totals(items)
    payload = chart_payload(items)
    recent = sorted(items, key=lambda e: (e.expense_date, e.id), reverse=True)[:10]
    users = get_users() if current_user.is_admin else []
    return render_template(
        "dashboard.html",
        totals=totals,
        expenses=recent,
        users=users,
        chart_payload=payload,
    )


@app.route("/expenses")
@login_required
def expenses():
    items = sorted(filter_expenses(current_expenses()), key=lambda e: (e.expense_date, e.id), reverse=True)
    page = parse_int(request.args.get("page"), 1) or 1
    pagination = manual_pagination(items, page, per_page=20)
    users = get_users() if current_user.is_admin else []
    return render_template("expenses.html", pagination=pagination, users=users)


@app.route("/expenses/new", methods=["GET", "POST"])
@login_required
def expense_new():
    if request.method == "POST":
        amount = parse_int(request.form.get("amount"))
        description = request.form.get("description", "").strip()
        category = request.form.get("category", "Other").strip() or "Other"
        expense_date = parse_date(request.form.get("expense_date"), date.today()) or date.today()
        if amount is None or amount <= 0 or not description:
            flash("Please enter a valid amount greater than zero and a description.", "error")
        elif category not in CATEGORIES:
            flash("Please select a valid category.", "error")
        else:
            insert_expense(
                user_id=current_user.id,
                amount=amount,
                description=description,
                category=category,
                expense_date=expense_date,
            )
            log_action(
                "expense_created",
                f"{current_user.username} added {money(amount)} expense: {description}",
                current_user.id,
            )
            flash("Expense added successfully.", "success")
            return redirect(url_for("dashboard"))
    return render_template("expense_form.html", expense=None)


@app.route("/expenses/<int:expense_id>/edit", methods=["GET", "POST"])
@login_required
def expense_edit(expense_id: int):
    expense = get_expense_by_id(expense_id)
    if expense is None:
        abort(404)
    if not current_user.is_admin and expense.user_id != current_user.id:
        abort(403)
    if request.method == "POST":
        amount = parse_int(request.form.get("amount"))
        description = request.form.get("description", "").strip()
        category = request.form.get("category", "Other").strip() or "Other"
        expense_date = parse_date(request.form.get("expense_date"), date.today()) or date.today()
        if amount is None or amount <= 0 or not description:
            flash("Please enter a valid amount greater than zero and a description.", "error")
        elif category not in CATEGORIES:
            flash("Please select a valid category.", "error")
        else:
            expense.amount = amount
            expense.description = description
            expense.category = category
            expense.expense_date = expense_date
            save_expense_record(expense)
            log_action(
                "expense_updated",
                f"{current_user.username} updated expense #{expense.id} to {money(amount)}",
                current_user.id,
            )
            flash("Expense updated successfully.", "success")
            return redirect(url_for("expenses"))
    expense.username = expense.username or current_user.username
    return render_template("expense_form.html", expense=expense)


@app.route("/expenses/<int:expense_id>/delete", methods=["POST"])
@login_required
def expense_delete(expense_id: int):
    expense = get_expense_by_id(expense_id)
    if expense is None:
        abort(404)
    if not current_user.is_admin and expense.user_id != current_user.id:
        abort(403)
    removed = delete_expense_record(expense_id)
    if removed is None:
        abort(404)
    log_action(
        "expense_deleted",
        f"{current_user.username} deleted {money(removed.amount)} expense: {removed.description}",
        current_user.id,
    )
    flash("Expense deleted successfully.", "success")
    return redirect(url_for("expenses"))


@app.route("/reports")
@login_required
def reports():
    items = filter_expenses(current_expenses())
    today = date.today()
    month_start = today.replace(day=1)
    next_month = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1)
    month_items = [expense for expense in items if month_start <= expense.expense_date < next_month]
    breakdown = monthly_breakdown(month_items)
    total = sum(breakdown.values())
    return render_template("reports.html", breakdown=breakdown, total=total)


@app.route("/export/xlsx")
@login_required
def export_xlsx():
    items = sorted(filter_expenses(current_expenses()), key=lambda e: (e.expense_date, e.id), reverse=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Expenses"
    sheet.append(["ID", "Date", "Category", "Description", "Amount", "User"])
    for row in items:
        sheet.append(
            [
                row.id,
                row.expense_date.isoformat(),
                row.category,
                row.description,
                row.amount,
                row.username,
            ]
        )
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="expenses.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/users", methods=["GET", "POST"])
@login_required
@admin_required
def admin_users():
    if request.method == "POST":
        action = request.form.get("action", "")
        user_id = parse_int(request.form.get("user_id"))
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if action == "create":
            if not username or not password:
                flash("Please provide both a username and a password.", "error")
            elif get_user_by_username(username) is not None:
                flash("That username is already taken. Please choose another.", "error")
            else:
                insert_user(
                    username,
                    generate_password_hash(password),
                    is_admin=parse_bool(request.form.get("is_admin")),
                    is_active_account=True,
                )
                log_action("user_created", f"Admin created user {username}", current_user.id)
                flash("User created successfully.", "success")
        elif action == "update" and user_id is not None:
            user = get_user_by_id(user_id)
            if user is None:
                flash("The selected user could not be found.", "error")
            else:
                if username and username != user.username and get_user_by_username(username) is not None:
                    flash("That username is already taken. Please choose another.", "error")
                elif user.is_admin and active_admin_count() <= 1 and (
                    not parse_bool(request.form.get("is_admin")) or not parse_bool(request.form.get("is_active_account"))
                ):
                    flash("At least one active administrator account must remain.", "error")
                else:
                    user.username = username or user.username
                    user.is_admin = parse_bool(request.form.get("is_admin"))
                    user.is_active_account = parse_bool(request.form.get("is_active_account"))
                    if password:
                        user.set_password(password)
                    save_user(user)
                    log_action("user_updated", f"Admin updated user {user.username}", current_user.id)
                    flash("User updated successfully.", "success")
        elif action == "toggle" and user_id is not None:
            user = get_user_by_id(user_id)
            if user and user.id != current_user.id:
                if user.is_admin and user.is_active_account and active_admin_count() <= 1:
                    flash("At least one active administrator account must remain.", "error")
                    return redirect(url_for("admin_users"))
                user.is_active_account = not user.is_active_account
                save_user(user)
                state = "activated" if user.is_active_account else "deactivated"
                log_action("user_toggled", f"Admin {state} user {user.username}", current_user.id)
                flash(f"User {state} successfully.", "success")
        return redirect(url_for("admin_users"))
    users = sorted(get_users(), key=lambda user: (not user.is_admin, user.username.lower()))
    return render_template("admin_users.html", users=users)


@app.route("/admin")
@login_required
@admin_required
def admin_home():
    return redirect(url_for("admin_users"))


@app.route("/admin/audit")
@login_required
@admin_required
def audit_log():
    return render_template("audit_log.html", entries=get_audit_entries(limit=200))


@app.route("/api/dashboard")
@login_required
def api_dashboard():
    items = filter_expenses(current_expenses())
    return jsonify(chart_payload(items))


@app.route("/manifest.json")
def manifest():
    return send_file(BASE_DIR / "static" / "manifest.json", mimetype="application/manifest+json")


@app.route("/sw.js")
def service_worker():
    return send_file(BASE_DIR / "static" / "sw.js", mimetype="application/javascript")


@app.errorhandler(403)
def forbidden(_):
    return render_template("error.html", code=403, message="You do not have permission to access this page."), 403


@app.errorhandler(404)
def not_found(_):
    return render_template("error.html", code=404, message="The page you are looking for could not be found."), 404


@app.errorhandler(500)
def server_error(_):
    return render_template(
        "error.html",
        code=500,
        message="Something went wrong on our end. Please try again in a few moments.",
    ), 500


def bootstrap() -> None:
    """Create tables and ensure the admin account exists. Runs once at startup."""
    with app.app_context():
        try:
            init_db()
            ensure_admin_account()
        except RuntimeError:
            raise


bootstrap()


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
    )
