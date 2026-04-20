#!/usr/bin/env python3
"""
setup_admin.py — One-time superadmin account setup.
Run this once after first install:
  cd automation-hub/backend
  ../.venv/bin/python3 setup_admin.py
"""

import sys
import getpass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from database import init_db, get_user_by_username, create_user
from werkzeug.security import generate_password_hash


def main():
    init_db()
    print('\n🔧  Payless Automation Hub — Superadmin Setup\n')

    username = input('Choose a username: ').strip()
    if not username:
        print('Username cannot be empty.')
        sys.exit(1)

    if get_user_by_username(username):
        print(f"User '{username}' already exists. Use the Admin panel to manage users.")
        sys.exit(1)

    password = getpass.getpass('Choose a password (min 6 chars): ')
    if len(password) < 6:
        print('Password must be at least 6 characters.')
        sys.exit(1)

    confirm = getpass.getpass('Confirm password: ')
    if password != confirm:
        print('Passwords do not match.')
        sys.exit(1)

    create_user(
        username      = username,
        password_hash = generate_password_hash(password),
        is_admin      = True,
        allowed_modules = ['*'],   # superadmin sees everything
    )

    print(f"\n✅  Superadmin '{username}' created successfully.")
    print('   You can now start the server and log in.\n')


if __name__ == '__main__':
    main()
