import unittest

from src.Clalit_Helper_Functions import choose_provider_index


class TestChooseProviderIndex(unittest.TestCase):

    def test_valid_israeli_ids_all_sums_1_to_13(self):
        """
        Each ID is 9 digits (Israeli-like),
        and the sum of digits modulo 13 covers 1–13.
        """

        test_cases = {
            "000000000": 0,
            "000000001": 1,   # sum = 1
            "000000002": 2,   # sum = 2
            "000000003": 3,   # sum = 3
            "000000004": 4,   # sum = 4
            "000000005": 5,   # sum = 5
            "000000006": 6,   # sum = 6
            "000000007": 7,   # sum = 7
            "000000008": 8,   # sum = 8
            "000000009": 9,   # sum = 9
            "000000019": 10,  # sum = 1 + 9 = 10
            "000000029": 11,  # sum = 2 + 9 = 11
            "000000039": 12,  # sum = 3 + 9 = 12
            "000000049": 13,  # sum = 4 + 9 = 13 → 13 % 13 = 0
        }

        for id_str, digit_sum in test_cases.items():
            with self.subTest(id=id_str):
                expected = digit_sum % 13
                self.assertEqual(choose_provider_index(id_str), expected)

    def test_string_input_required(self):
        """Ensure string input works correctly"""
        self.assertEqual(choose_provider_index("123456789"), (1+2+3+4+5+6+7+8+9) % 13)

    def test_leading_zeros(self):
        """Leading zeros should not affect the result"""
        self.assertEqual(choose_provider_index("000123456"), (1+2+3+4+5+6) % 13)

    def test_large_israeli_like_id(self):
        """Typical 9-digit Israeli ID"""
        self.assertEqual(choose_provider_index("987654321"), (9+8+7+6+5+4+3+2+1) % 13)

    def test_custom_providers_count(self):
        """Check modulo behavior with different providers_count"""
        self.assertEqual(choose_provider_index("123", providers_count=5), (1+2+3) % 5)

    def test_zero_id(self):
        """Edge case: ID with all zeros"""
        self.assertEqual(choose_provider_index("000000000"), 0)


if __name__ == '__main__':
    unittest.main()
