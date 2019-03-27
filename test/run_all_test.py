import unittest

if __name__ == '__main__':
    suite = unittest.TestSuite()
    all_cases = unittest.defaultTestLoader.discover('.', '*Test.py')
    for case in all_cases:
        suite.addTests(case)
    runner = unittest.TextTestRunner()
    runner.run(suite)
