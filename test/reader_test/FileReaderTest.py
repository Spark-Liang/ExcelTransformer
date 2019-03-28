import json
import unittest
from os import path

import pandas as pd
from pandas import DataFrame

from common import Template
from reader import FileReader as SUT


class TestReadExcel(unittest.TestCase):

    def test_read_single_mapping(self):
        # given
        test_name = "ReadSingleMapping"
        file_type = "excel"
        expected_set_of_data_frame = read_expected_result(test_name, file_type)
        source_path = get_test_data_source_path(test_name, file_type)
        template_path = get_test_template_path(test_name, file_type)
        config = read_reader_config(test_name, file_type)
        sut = SUT(Template(template_path), config)

        # when
        result = sut.read(source_path)

        # then
        self.assertEqual(
            expected_set_of_data_frame, set(result)
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
    :rtype:set[DataFrame] or DataFrame
    """
    data_file_path = path.join(
        get_test_data_prefix(test_name, file_type),
        "Result.{}".format(
            "xlsx" if file_type == "excel" else "csv"
        )
    )

    if file_type == "excel":
        return pd.read_excel(
            data_file_path, sheet_name=None
        )
    else:
        return pd.read_csv(data_file_path)
