import unittest
from os import path

from test_util import MessageUtil
from xl_transform.common import Cell
from xl_transform.common import ParseTemplate as SUT
from xl_transform.common.ParseTemplate import TemplateInfoItem, CellTemplateInfoItem, row, column


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
        result = SUT.parse_excel_template(template_path)[0]

        # then
        result_item = result[0]
        self.assertEqual("TestOut", result_item.sheet_name)
        expected_cell = Cell(1, 1)
        self.assertEqual(
            expected_cell,
            result_item.top_left_point,
            "expected:{}\nresult:{}".format(expected_cell, result_item.top_left_point)
        )
        self.assertEqual("TBL_1", result_item.mapping_name)
        self.assertEqual(("1", "2", "3", "4"), result_item.headers)
        self.assertEqual(row, result_item.header_direction)

    def test_can_parse_two_mapping(self):
        # given
        template_path = self.get_test_file_path("TemplateWithTwoMapping.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)[0]

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(1, 1),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                "TestOut", Cell(1, 5),
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
        result = SUT.parse_excel_template(template_path)[0]

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(1, 1),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                "TestOut", Cell(8, 1),
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
        result = SUT.parse_excel_template(template_path)[0]

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(1, 1),
                "TBL_1", ["1", "2", "3", "4"], row
            ),
            TemplateInfoItem(
                "TestOut", Cell(8, 1),
                "TBL_2", ["A", "B", "C", "D"], column
            ),
            TemplateInfoItem(
                "TestSecondSheet", Cell(1, 1),
                "TBL_3", ["a", "b", "c", "d"], column
            ),
            TemplateInfoItem(
                "TestSecondSheet", Cell(10, 1),
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

    def test_header_direction_is_None_when_header_is_single_cell(self):
        # given
        template_path = self.get_test_file_path("TemplateWithSingleCell.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)[0]

        # then
        expected_items = {
            TemplateInfoItem(
                "TestOut", Cell(2, 2),
                "TBL_1", ["1"], None
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )

    def test_parse_single_cell_extractor(self):
        # given
        template_path = self.get_test_file_path("TemplateWithSingleCellExtract.xlsx")

        # when
        result = SUT.parse_excel_template(template_path)[1]

        # then
        expected_items = {
            CellTemplateInfoItem(
                "TestSingleCellExtract",
                "test_param",
                Cell(6, 3)
            )
        }
        self.assertEqual(
            expected_items, set(result),
            "expected_items: {};\nresult: {}".format(
                MessageUtil.format_collection(expected_items),
                MessageUtil.format_collection(result)
            )
        )
