import json
import unittest
from os import path

import pandas as pd
from pandas import DataFrame

from common import TemplateInfoItem, Cell, row, column
from reader import DataFrameReader as SUT


class TestReadExcel(unittest.TestCase):

    def test_read_skip_header_row(self):
        # given
        test_name = "ReadSkipHeaderRow"
        file_type = "excel"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            "Data", Cell(2, 2), "TBL_1",
            ["A", "B", "C", "D"], row
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def test_read_with_vertical_header_and_skip_header_row(self):
        # given
        test_name = "ReadWithVerticalHeaderSkipHeaderRow"
        file_type = "excel"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            "Data", Cell(2, 2), "TBL_1",
            ["A", "B", "C", "D"], column
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def test_auto_assign_header(self):
        # given
        test_name = "ReadHeaderAndAutoAssignToDataFrame"
        file_type = "excel"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            "Data", Cell(2, 2), "TBL_1",
            ["_", "_", "_", "_"], row
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def test_can_limit_rows_to_read(self):
        # given
        test_name = "ReadWithFixedNumOfRow"
        file_type = "excel"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            "Data", Cell(2, 2), "TBL_1",
            ["A", "B", "C", "D"], row
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result
    ):
        """

        :param DataFrame expected_data_frame:
        :param DataFrame result:
        :return:
        """
        self.assertEqual(
            expected_data_frame.shape, result.shape
        )
        for expected_header, result_header in zip(expected_data_frame.columns.values, result.columns.values):
            self.assertEqual(expected_header, result_header)
        max_x, max_y = expected_data_frame.shape
        for x in range(0, max_x):
            for y in range(0, max_y):
                self.assertEqual(
                    expected_data_frame.iloc[x, y],
                    result.iloc[x, y]
                )


# noinspection PyTypeChecker
class TestReadCSV(unittest.TestCase):

    def test_read_skip_header_row(self):
        # given
        test_name = "ReadSkipHeaderRow"
        file_type = "csv"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            None, Cell(0, 0), "TBL_1",
            ["A", "B", "C", "D"], row
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def test_read_with_vertical_header_and_skip_header_row(self):
        # given
        test_name = "ReadWithVerticalHeaderSkipHeaderRow"
        file_type = "csv"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            None, Cell(0, 0), "TBL_1",
            ["A", "B", "C", "D"], column
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def test_auto_assign_header(self):
        # given
        test_name = "ReadHeaderAndAutoAssignToDataFrame"
        file_type = "csv"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            None, Cell(0, 0), "TBL_1",
            ["_", "_", "_", "_"], row
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )

    def test_can_limit_rows_to_read(self):
        # given
        test_name = "ReadWithFixedNumOfRow"
        file_type = "csv"
        expected_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        info_item = TemplateInfoItem(
            None, Cell(0, 0), "TBL_1",
            ["A", "B", "C", "D"], row
        )
        sut = SUT(info_item, config)

        # when
        result = sut.read(source_path)

        # then

        assert_data_frame_with_same_header_and_data(
            self, expected_data_frame, result[1]
        )


def assert_data_frame_with_same_header_and_data(
        test_case, expected_data_frame, result
):
    """

    :param test_case:
    :param DataFrame expected_data_frame:
    :param DataFrame result:
    :return:
    """
    test_case.assertEqual(
        expected_data_frame.shape, result.shape
    )
    for expected_header, result_header in zip(expected_data_frame.columns.values, result.columns.values):
        test_case.assertEqual(expected_header, result_header)
    max_x, max_y = expected_data_frame.shape
    for x in range(0, max_x):
        for y in range(0, max_y):
            test_case.assertEqual(
                expected_data_frame.iloc[x, y],
                result.iloc[x, y]
            )


def get_test_data_prefix(test_name, file_type):
    return path.join(
        path.dirname(__file__), "test_data", "DataFrameReader", file_type, test_name
    )


def get_test_data_source_path(testname, file_type):
    return path.join(
        get_test_data_prefix(testname, file_type),
        "Source.{}".format(
            "xlsx" if file_type == "excel" else "csv"
        )
    )


def read_expected_result(test_name, file_type):
    """

    :param file_type:
    :param test_name:
    :return:
    """
    data_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "Result.{}".format(
            "xlsx" if file_type == "excel" else "csv"
        )
    )

    if file_type == "excel":
        return pd.read_excel(
            data_file_path
        )
    else:
        return pd.read_csv(data_file_path)


def read_reader_config(test_name, file_type):
    config_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "config.json"
    )
    with open(config_file_path, "r") as json_file:
        return json.loads(
            "".join(json_file.readlines())
        )
