import subprocess
import sys
import unittest


class TestCLI(unittest.TestCase):
    def test_cli_help(self):
        res = subprocess.run(
            [sys.executable, "-m", "xlsform_to_registro", "--help"],
            capture_output=True,
            text=True,
        )
        self.assertEqual(res.returncode, 0)
        self.assertIn("Generador", res.stdout)
        self.assertIn("--header-color", res.stdout)


if __name__ == "__main__":
    unittest.main()
