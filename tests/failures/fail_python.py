"""Deliberately failing stdlib test for AC-009; excluded from normal discovery."""

import unittest


class KnownFailure(unittest.TestCase):
    def test_known_failure(self):
        self.fail("Intentional synthetic Python failure for runner propagation")


if __name__ == "__main__":
    unittest.main(verbosity=2)
