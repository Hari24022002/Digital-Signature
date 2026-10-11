"""Create a Certifier account for a local e-Hastakshar installation.

Run from the project directory with the project virtual environment activated:
    python create_certifier.py
"""
import getpass
import json
import os
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.json")


def main():
    username = input("Certifier display name: ").strip()
    pbno = input("Certifier employee/PB number: ").strip()
    password = getpass.getpass("Set certifier password: ")
    confirm = getpass.getpass("Confirm password: ")

    if not username or not pbno or not password:
        raise SystemExit("Name, PB number, and password are required.")
    if password != confirm:
        raise SystemExit("Passwords do not match.")

    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as file:
                users = json.load(file)
        except (json.JSONDecodeError, OSError) as exc:
            raise SystemExit(f"Could not read {USERS_FILE}: {exc}")
    else:
        users = []

    if any(user.get("pbno") == pbno for user in users):
        raise SystemExit("A user with that PB number already exists.")

    users.append({
        "username": username,
        "pbno": pbno,
        "password": generate_password_hash(password),
        "role": "Certifier",
        "stamp": ""
    })
    with open(USERS_FILE, "w", encoding="utf-8") as file:
        json.dump(users, file, indent=4)
    print("Certifier account created. Keep users.json private.")


if __name__ == "__main__":
    main()
