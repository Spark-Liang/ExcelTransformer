import unittest
from os import path

from common import ParseTemplate as SUT
from common.ParseTemplate import TemplateInfoItem, row, column
from common.model import Cell
from test_util import MessageUtil


class TestParseExcelTemplate(unittest.TestCase):

    def setUp(self):
        self.test_data_path = path.join(path.dirname(__file__), "test_data")

    def get_test_file_path(self, test_file_name):
        """
        :param str test_file_name:
        :return:
        """
        return path.join(self.test_data_path, test_file_name)

    def test_can_parse_single_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithSingleMapping.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)

        # then
        result_item = result[0]
        self.assertEqual("TestOut", result_item.sheet_name)
        self.assertEqual(Cell(0, 0), result_item.top_left_point)
        self.assertEqual("TBL_1", result_item.mapping_name)
        self.assertEqual(("1", "2", "3", "4"), result_item.headers)
        self.assertEqual(row, result_item.header_direction)

    def test_can_parse_two_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithTwoMapping.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(0, 0),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                "TestOut", Cell(0, 4),
                "TBL_2", ["A", "B", "C", "D"], row
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )

    def test_can_parse_with_vertical_and_horizontal_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithHorizontalAndVeritcalMapping.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(0, 0),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                "TestOut", Cell(7, 0),
                "TBL_2", ["A", "B", "C", "D"], column
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )

    def test_can_parse_with_two_sheet(self):
        # given
        template_path = self.get_test_file_path("TemplateWithTwoSheet.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(0, 0),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                "TestOut", Cell(7, 0),
                "TBL_2", ["A", "B", "C", "D"], column
            ),
            TemplateInfoItem(
                "TestSecondSheet", Cell(0, 0),
                "TBL_3", ["a", "b", "c", "d"], column
            ),
            TemplateInfoItem(
                "TestSecondSheet", Cell(9, 0),
                "TBL_4", ["abc", "bcd", "cde", "def"], row
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )


# noinspection PyTypeChecker
class TestParseCSVTemplate(unittest.TestCase):

    def setUp(self):
        self.test_data_path = path.join(path.dirname(__file__), "test_data")

    def get_test_file_path(self, test_file_name):
        """
        :param str test_file_name:
        :return:
        """
        return path.join(self.test_data_path, test_file_name)

    def test_can_parse_single_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithSingleMapping.csv")

        # when
        result = SUT.parse_csv_template(template_path)

        # then
        result_item = result[0]
        self.assertEqual(None, result_item.sheet_name)
        self.assertEqual(Cell(0, 0), result_item.top_left_point)
        self.assertEqual("TBL_1", result_item.mapping_name)
        self.assertEqual(("1", "2", "3", "4"), result_item.headers)
        self.assertEqual(row, result_item.header_direction)

    def test_can_parse_two_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithTwoMapping.csv")

        # when
        result = SUT.parse_csv_template(template_path)

        # then
        expected_items = {
            TemplateInfoItem(
                None, Cell(0, 0),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                None, Cell(0, 4),
                "TBL_2", ["A", "B", "C", "D"], row
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )

    def test_can_parse_with_vertical_and_horizontal_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithHorizontalAndVeritcalMapping.csv")

        # when
        result = SUT.parse_csv_template(template_path)

        # then
        expected_items = {
            TemplateInfoItem(
                None, Cell(0, 0),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                None, Cell(7, 0),
                "TBL_2", ["A", "B", "C", "D"], column
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )
