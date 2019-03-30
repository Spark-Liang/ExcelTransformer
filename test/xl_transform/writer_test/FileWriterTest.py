import json
import os
import unittest
from os import path

import pandas as pd

from test_util.DataFrameAssertUtil import assert_data_in_excel_is_equal, assert_data_in_data_frame_is_equal
from xl_transform.common import Template
from xl_transform.writer import FileWriter as SUT


class TestWriteExcel(unittest.TestCase):

    def setUp(self):
        self.__test_target_path = "init"

    def tearDown(self):
        if path.exists(self.__test_target_path):
            os.remove(self.__test_target_path)
            self.__test_target_path = None

    def test_write_with_horizontal_and_vertical_mapping(self):
        # given
        test_name = "WriteHorizontalAndVerticalMapping"
        file_type = "excel"
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        data_frame_dict = get_test_data(test_name, file_type)
        sut = SUT(Template(template_path), config)

        # when
        sut.write(data_frame_dict, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, test_name)

    def test_write_data_to_two_sheet(self):
        # given
        test_name = "WriteDataIntoTwoSheet"
        file_type = "excel"
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        data_frame_dict = get_test_data(test_name, file_type)
        sut = SUT(Template(template_path), config)

        # when
        sut.write(data_frame_dict, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, test_name)

    def test_raise_exception_when_area_to_write_is_contacted(self):
        # given
        test_name = "WriteFailedWhenTwoMappingAreaHasContacted"
        file_type = "excel"
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        data_frame_dict = get_test_data(test_name, file_type)
        sut = SUT(Template(template_path), config)

        # when
        def func():
            sut.write(data_frame_dict, self.__test_target_path)

        self.assertRaises(Exception, func)

    def assert_expected_equal_to_result(self, file_type, test_name):
        self.assertTrue(
            path.exists(self.__test_target_path)
        )
        assert_data_in_excel_is_equal(
            self,
            get_expected_target_file_path(test_name, file_type),
            self.__test_target_path
        )


class TestWriteToCSV(unittest.TestCase):

    def test_write_data_into_two_mapping(self):
        # given
        test_name = "WriteDataIntoTwoMapping"
        file_type = "csv"
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        data_frame_dict = get_test_data(test_name, file_type)
        sut = SUT(Template(template_path), config)

        # when
        sut.write(data_frame_dict, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, test_name)

    def assert_expected_equal_to_result(self, file_type, test_name):
        expected_df = pd.read_csv(
            get_expected_target_file_path(test_name, file_type),
            header=None
        )
        result_df = pd.read_csv(
            self.__test_target_path, header=None
        )
        assert_data_in_data_frame_is_equal(
            self, expected_df, result_df
        )


def get_test_data_prefix(test_name, file_type):
    return path.join(
        path.dirname(__file__), "test_data", "FileWriter", file_type, test_name
    )


def get_test_data(test_name, file_type):
    file_path = path.join(
        get_test_data_prefix(test_name, file_type), "Data.xlsx"
    )
    return pd.read_excel(
        file_path, sheet_name=None
    )


def get_test_template_path(test_name, file_type):
    return path.join(
        get_test_data_prefix(test_name, file_type),
        "Template.{}".format(
            "xlsx" if file_type == "excel" else "csv"
        )
    )


def get_test_target_file_path(test_name, file_type):
    suffix = "xlsx" if file_type == "excel" else "csv"
    test_target_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "{}.{}".format("TmpOut", suffix)
    )
    if path.exists(test_target_file_path):
        os.remove(test_target_file_path)
    return test_target_file_path


def get_expected_target_file_path(test_name, file_type):
    suffix = "xlsx" if file_type == "excel" else "csv"
    return path.join(
        get_test_data_prefix(test_name, file_type),
        "Result." + suffix
    )


def read_writer_config(test_name, file_type):
    config_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "config.json"
    )
    with open(config_file_path, "r") as json_file:
        return json.loads(
            "".join(json_file.readlines())
        )
