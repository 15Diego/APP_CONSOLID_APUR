import unittest

from supabase_repository import _validate_connection_value


class SupabaseSecretValidationTests(unittest.TestCase):
    def test_accepts_ascii_connection_values(self):
        _validate_connection_value("SUPABASE_SERVICE_ROLE_KEY", "eyJhbGciOiJIUzI1NiJ9.example")

    def test_rejects_unicode_with_clear_message(self):
        with self.assertRaisesRegex(RuntimeError, "Copie apenas o valor real da chave"):
            _validate_connection_value("SUPABASE_SERVICE_ROLE_KEY", "chave → exemplo")


if __name__ == "__main__":
    unittest.main()
