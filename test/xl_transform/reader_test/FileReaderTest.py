import json
import unittest
from os import path

import pandas as pd

from test_util.DataFrameAssertUtil import assert_data_frame_with_same_header_and_data
from xl_transform.common import Template
from xl_transform.reader import FileReader as SUT


class TestReadExcel(unittest.TestCase):

    def test_read_single_mapping(self):
        # given
        test_name = "ReadSingleMapping"
        file_type = "excel"

        # when, then
        self.read_and_check_data(file_type, test_name)

    def test_read_data_from_two_sheet(self):
        # given
        test_name = "ReadTwoSheet"
        file_type = "excel"

        # when, then
        self.read_and_check_data(file_type, test_name)

    def test_read_with_horizontal_and_vertical_header(self):
        # given
        test_name = "ReadHorizontalAndVerticalData"
        file_type = "excel"

        # when, then
        self.read_and_check_data(file_type, test_name)

    def test_read_fail_when_two_mapping_with_same_name_in_the_template(self):
        # given
        test_name = "ReadFailedWhenHaveTwoMappingWithSameName"
        file_type = "excel"

        template_path = get_test_template_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)

        # when
        def func():
            sut = SUT(Template(template_path), config)

        self.assertRaises(Exception, func)

    def read_and_check_data(self, file_type, test_name):
        expected_dict_of_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        sut = SUT(Template(template_path), config)
        # when
        result = sut.read(source_path)
        # then
        assert_expected_equal_to_result(self, expected_dict_of_data_frame, result)


class TestReadCSV(unittest.TestCase):

    def test_read_single_mapping(self):
        # given
        test_name = "ReadSingleMapping"
        file_type = "csv"
        test_mapping_list = ["TBL_1"]

        source_path = get_test_data_source_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        expected_dict_of_data_frame = TestReadCSV.read_expected_result_from_csv(
            test_name, test_mapping_list
        )
        sut = SUT(Template(template_path), config)

        # when
        result = sut.read(source_path)
        # then
        assert_expected_equal_to_result(self, expected_dict_of_data_frame, result)

    def test_read_two_mapping(self):
        # given
        test_name = "ReadTwoMapping"
        file_type = "csv"
        test_mapping_list = ["TBL_1", "TBL_2"]

        source_path = get_test_data_source_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        expected_dict_of_data_frame = TestReadCSV.read_expected_result_from_csv(
            test_name, test_mapping_list
        )
        sut = SUT(Template(template_path), config)

        # when
        result = sut.read(source_path)
        # then
        assert_expected_equal_to_result(self, expected_dict_of_data_frame, result)

    def test_read_fail_when_two_mapping_with_same_name_in_the_template(self):
        # given
        test_name = "ReadFailedWhenHaveTwoMappingWithSameName"
        file_type = "csv"

        template_path = get_test_template_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)

        # when
        def func():
            sut = SUT(Template(template_path), config)

        self.assertRaises(Exception, func)

    @staticmethod
    def read_expected_result_from_csv(test_name, mapping_name_list):
        expected_results = {}
        bash_path = get_test_data_prefix(test_name, "csv")
        for mapping_name in mapping_name_list:
            file_path = path.join(bash_path, "Result_{}.csv".format(mapping_name))
            expected_results[mapping_name] = pd.read_csv(
                file_path,
                # add the type hint just for the reason that,
                # we can compare the data by invoke the "equal" method.
                dtype=str
            )

        return expected_results


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


def get_test_data_prefix(test_name, file_type):
    return path.join(
        path.dirname(__file__), "test_data", "FileReader", file_type, test_name
    )


def get_test_data_source_path(testname, file_type):
    return path.join(
        get_test_data_prefix(testname, file_type),
        "Source.{}".format(
            "xlsx" if file_type == "excel" else "csv"
        )
    )


def get_test_template_path(test_name, file_type):
    return path.join(
        get_test_data_prefix(test_name, file_type),
        "Template.{}".format(
            "xlsx" if file_type == "excel" else "csv"
        )
    )


def read_reader_config(test_name, file_type):
    config_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "config.json"
    )
    with open(config_file_path, "r") as json_file:
        return json.loads(
            "".join(json_file.readlines())
        )


def read_expected_result(test_name, file_type):
    """

    :param file_type:
    :param test_name:
    :return:
    """
    data_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "Result.xlsx"
    )
    return pd.read_excel(
        data_file_path, sheet_name=None
    )
