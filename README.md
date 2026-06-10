# Expense Tracker

Simple Flask expense tracker backed by CSV files instead of a database server.

## Features

- User login and admin login
- Admin can create, edit, deactivate, and reset user passwords
- Expense CRUD with amount, quantity, description, date, and category
- Dashboard totals for today, this week, this month, and lifetime
- Search and filters by date, category, amount, and user
- Monthly category report
- CSV and Excel export
- Audit log for admin actions
- Mobile-friendly UI with PWA support

## CSV storage

Data is stored in `data/users.csv`, `data/expenses.csv`, and `data/audit.csv`.

## Local run

```bash
pip install -r requirements.txt
python app.py
```

## Render deployment

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`
- Environment variables:
  - `SECRET_KEY`
  - `ADMIN_USERNAME` optional
  - `ADMIN_PASSWORD` optional
  - `DATA_DIR` optional, use this for a persistent Render disk path
  - `REMEMBER_DAYS` optional, defaults to `30`

- Users are created from the admin panel at `/admin`
- Login sessions are remembered for 30 days by default

If you want the CSV files to persist on Render, attach a persistent disk and set `DATA_DIR` to the mounted path.
