import unittest
from os import path

import pandas as pd

from test_util.DataFrameAssertUtil import assert_data_frame_with_same_header_and_data
from xl_transform.reader import FileReader as SUT


class TestRead(unittest.TestCase):

    def test_read_single_mapping(self):
        # given
        test_name = "ReadSingleMapping"

        # when, then
        self.read_and_check_data(test_name)

    def test_read_data_from_two_sheet(self):
        # given
        test_name = "ReadTwoSheet"

        # when, then
        self.read_and_check_data(test_name)

    def test_read_with_horizontal_and_vertical_header_and_without_skip_header(self):
        # given
        test_name = "ReadHorizontalAndVerticalDataAndWithoutSkipHeader"

        # when, then
        self.read_and_check_data(test_name)

    def test_read_with_rows_limit(self):
        # given
        test_name = "ReadWithRowsLimit"

        # when, then
        self.read_and_check_data(test_name)

    def test_read_single_row_and_single_column(self):
        # given
        test_name = "ReadSingleRowAndSingleColumn"

        # when, then
        self.read_and_check_data(test_name)

    def test_read_fail_when_two_mapping_with_same_name_in_the_template(self):
        # given
        test_name = "ReadFailedWhenHaveTwoMappingWithSameName"

        template_path = get_test_template_path(test_name)
        config = get_reader_config_path(test_name)
        source_path = get_test_data_source_path(test_name)

        # when
        def func():
            SUT.read(source_path, template_path, config)

        self.assertRaises(Exception, func)

    def read_and_check_data(self, test_name):
        expected_dict_of_data_frame = read_expected_result(test_name)
        source_path = get_test_data_source_path(test_name)
        template_path = get_test_template_path(test_name)
        config_path = get_reader_config_path(test_name)
        # when
        result = SUT.read(source_path, template_path, config_path)
        # then
        assert_expected_equal_to_result(self, expected_dict_of_data_frame, result)


def assert_expected_equal_to_result(test_case, expected_dict_of_data_frame, result):
    test_case.assertSetEqual(
        set(expected_dict_of_data_frame.keys()),
        set(result.keys())
    )
    for mapping_name in expected_dict_of_data_frame.keys():
        assert_data_frame_with_same_header_and_data(
            test_case,
            expected_dict_of_data_frame[mapping_name],
            result[mapping_name]
        )


def get_test_data_prefix(test_name):
    return path.join(
        path.dirname(__file__), "test_data", "FileReader", test_name
    )


def get_test_data_source_path(testname):
    return path.join(
        get_test_data_prefix(testname),
        "Source.xlsx"
    )


def get_test_template_path(test_name):
    return path.join(
        get_test_data_prefix(test_name),
        "Template.xlsx"
    )


def get_reader_config_path(test_name):
    return path.join(
        get_test_data_prefix(test_name),
        "config.json"
    )


def read_expected_result(test_name):
    """

    :param test_name:
    :return:
    """
    data_file_path = path.join(
        get_test_data_prefix(test_name),
        "Result.xlsx"
    )
    return pd.read_excel(
        data_file_path, sheet_name=None
    )
