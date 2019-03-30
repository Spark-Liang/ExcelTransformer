import unittest
from os import path

from test_util import HTMLTestRunner

if __name__ == '__main__':
    suite = unittest.TestSuite()
    all_cases = unittest.defaultTestLoader.discover(
        path.join(path.dirname(__file__), "xl_transform"),
        '*Test.py'
    )
    for case in all_cases:
        suite.addTests(case)

    report_file_path = path.join(path.dirname(__file__), "htmlreport.html")
    with open(report_file_path, "wb") as f:
        runner = HTMLTestRunner(
            stream=f,
            title="ExcelTransformer Unittest Report"
        )
        runner.run(suite)
        unittest.main()
    # unittest.TextTestRunner().run(suite)
