import json
import os
import tempfile
import unittest

import app as signature_app


class DigitalSignatureSmokeTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_paths = {
            "DOCS_FILE": signature_app.DOCS_FILE,
            "HISTORY_FILE": signature_app.HISTORY_FILE,
            "USERS_FILE": signature_app.USERS_FILE,
            "UPLOAD_FOLDER": signature_app.UPLOAD_FOLDER,
        }
        signature_app.DOCS_FILE = os.path.join(self.temp_dir.name, "documents.json")
        signature_app.HISTORY_FILE = os.path.join(self.temp_dir.name, "history.json")
        signature_app.USERS_FILE = os.path.join(self.temp_dir.name, "users.json")
        signature_app.UPLOAD_FOLDER = os.path.join(self.temp_dir.name, "upload")
        os.makedirs(signature_app.UPLOAD_FOLDER, exist_ok=True)
        signature_app.app.config.update(TESTING=True)
        self.client = signature_app.app.test_client()

    def tearDown(self):
        for key, value in self.original_paths.items():
            setattr(signature_app, key, value)
        self.temp_dir.cleanup()

    def test_home_page_renders(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)

    def test_registration_hashes_password_and_login_works(self):
        response = self.client.post("/register", data={
            "username": "prepared-user",
            "pbno": "PB1001",
            "password": "test-password",
            "role": "Prepared",
        })
        self.assertEqual(response.status_code, 302)
        with open(signature_app.USERS_FILE, "r", encoding="utf-8") as file:
            users = json.load(file)
        self.assertEqual(len(users), 1)
        self.assertNotEqual(users[0]["password"], "test-password")

        login = self.client.post("/login", data={
            "pbno": "PB1001",
            "password": "test-password",
        })
        self.assertEqual(login.status_code, 302)
        dashboard = self.client.get("/dashboard")
        self.assertEqual(dashboard.status_code, 200)

    def test_protected_routes_reject_anonymous_requests(self):
        action = self.client.post("/action/example/Approved")
        self.assertEqual(action.status_code, 401)
        uploaded = self.client.get("/uploads/anything.pdf")
        self.assertEqual(uploaded.status_code, 401)

    def test_prepared_user_cannot_approve_documents(self):
        self.client.post("/register", data={
            "username": "prepared-user",
            "pbno": "PB1002",
            "password": "test-password",
            "role": "Prepared",
        })
        self.client.post("/login", data={
            "pbno": "PB1002",
            "password": "test-password",
        })
        action = self.client.post("/action/example/Approved")
        self.assertEqual(action.status_code, 403)


if __name__ == "__main__":
    unittest.main()
