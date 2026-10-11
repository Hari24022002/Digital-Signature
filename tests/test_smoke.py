import unittest

import app as signature_app


class DigitalSignatureSmokeTests(unittest.TestCase):
    def setUp(self):
        signature_app.app.config.update(TESTING=True)
        self.client = signature_app.app.test_client()

    def test_home_page_renders(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)

    def test_protected_routes_reject_anonymous_requests(self):
        action = self.client.post("/action/example/Approved")
        self.assertEqual(action.status_code, 401)
        uploaded = self.client.get("/uploads/anything.pdf")
        self.assertEqual(uploaded.status_code, 401)


if __name__ == "__main__":
    unittest.main()
