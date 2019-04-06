import unittest
from datetime import datetime
from decimal import Decimal
from os import path

from pandas import DataFrame

from test_util.DataFrameAssertUtil import assert_data_in_data_frame_is_equal
from xl_transform.reader import ExcelDataExtractor as SUT


class ExcelExtractTest(unittest.TestCase):

    def test_can_get_right_data_frame(self):
        # given
        expected_data_frame = DataFrame({
            "str": ["123", "ABC", "BCD"],
            "decimal": [1, 1.1, -1],
            "date": [datetime.strptime(val, "%Y/%m/%d %H:%M:%S")
                     for val in ["2019/12/31 23:59:22", "2019/11/30 23:59:22", "2019/01/31 23:59:22"]]
        })
        source_data_path = path.join(path.dirname(__file__), "test_data", "TestExcelDataExtractData.xlsx")

        # when
        result = SUT.extract_data_frame(
            source_data_path, sheet_name="horizontal",
            header_list=["str", "decimal", "date"],
            start_row_idx=2,
            dtype={
                "str": str, "decimal": Decimal
            }
        )

        # then
        assert_data_in_data_frame_is_equal(
            self,
            expected_data_frame, result
        )
