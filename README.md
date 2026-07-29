# XPLMS

XPLMS is a role-based Experience Point Learning Management System built with
Streamlit and Supabase.

## Features

- Student and administrator workspaces
- Student progress and published-results tracking
- Student CRUD and guarded CSV/Excel bulk imports
- XP, badges, performance analytics and leaderboard
- Study-material library backed by private Supabase Storage
- Supabase Auth and row-level security
- Responsive student experience for mobile devices

## Setup

1. Create a Python virtual environment and install `requirements.txt`.
2. Keep `SUPABASE_URL` and `SUPABASE_KEY` in `.streamlit/secrets.toml`.
3. Run `supabase_schema.sql` once in the Supabase SQL Editor.
4. Add users with Supabase Authentication, then add their profile and role.
5. Start the app with `streamlit run app.py`.

The sign-in screen includes safe demo workspaces for reviewing all three roles.
Demo actions do not write to Supabase.
