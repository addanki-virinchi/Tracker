# Expense Tracker

Simple Flask expense tracker backed by Firebase Realtime Database.

## Features

- User login and admin login
- Admin can create, edit, deactivate, and reset user passwords
- Expense CRUD with amount, quantity, description, date, and category
- Dashboard totals for today, this week, this month, and lifetime
- Search and filters by date, category, amount, and user
- Monthly category report
- Excel export
- Audit log for admin actions
- Mobile-friendly UI with PWA support

## Firebase storage

Data is stored in Firebase Realtime Database under `users`, `expenses`, `audit`, and `meta/counters`.

## Local run

```bash
pip install -r requirements.txt
python app.py
```

## Firebase setup

- Set `FIREBASE_DATABASE_URL` to your Realtime Database URL.
- Set `FIREBASE_CREDENTIALS_PATH` to the Firebase service-account JSON file, or
  set `FIREBASE_CREDENTIALS_JSON` to the JSON contents directly.
- `ADMIN_USERNAME` and `ADMIN_PASSWORD` control the bootstrap admin account.

## Render deployment

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`
- Environment variables:
  - `SECRET_KEY`
  - `FIREBASE_DATABASE_URL`
  - `FIREBASE_CREDENTIALS_JSON` or `FIREBASE_CREDENTIALS_PATH`
  - `ADMIN_USERNAME` optional
  - `ADMIN_PASSWORD` optional
  - `REMEMBER_DAYS` optional, defaults to `30`

- Users are created from the admin panel at `/admin`
- Login sessions are remembered for 30 days by default
