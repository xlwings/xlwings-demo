import unittest


def mysum(a, b):
    return a + b


class TestMySum(unittest.TestCase):

    def test_postive(self):
        self.assertEqual(mysum(1, 2), 3)


if __name__ == '__main__':
    unittest.main()
