import json
import os
import time
import unittest
from decimal import Decimal
from os import path

import pandas as pd

from test_util.DataFrameAssertUtil import assert_data_in_excel_is_equal
from xl_transform.common import TemplateInfoItem, row, column
from xl_transform.common.model import Cell
from xl_transform.writer import ExcelDataFrameWriter as SUT


class TestWrite(unittest.TestCase):

    def setUp(self):
        self.__test_target_path = "init"

    def tearDown(self):
        if path.exists(self.__test_target_path):
            os.remove(self.__test_target_path)
            self.__test_target_path = None

    def test_write_with_format(self):
        # given
        test_name = "WriteWithFormat"
        file_type = "excel"
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        target_sheet_name = "Target"
        info_item = TemplateInfoItem(
            target_sheet_name, Cell(0, 0), "TBL_1",
            ["str", "int", "date_1", "date_2", "float", "decimal"], row
        )
        date_format = "%Y/%m/%d %H:%M:%S"
        data_frame = pd.DataFrame({
            "str": ["10", "0.029576325", "0.333226408"],
            "int": ["20", "-10", "0"],
            "date_1": [time.strptime(str_val, date_format) for str_val in
                       ["2019/01/30 00:00:00", "2019/02/28 16:58:24", "2019/12/30 08:20:47"]],
            "date_2": [time.strptime(str_val, date_format) for str_val in
                       ["2019/01/30 00:00:00", "2019/02/28 16:58:24", "2019/12/30 08:20:47"]],
            "float": [6.144880061, -1.195569306, 2.847225604],
            "decimal": [Decimal.from_float(x) for x in [6.144880061, -1.195569306, 2.847225604]]
        })
        sut = SUT(info_item, config)

        # when
        sut.write(data_frame, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, target_sheet_name, test_name)

    def test_write_with_single_mapping_and_with_projection(self):
        # given
        test_name = "WriteSingleMappingWithProjection"
        file_type = "excel"
        data_frame = get_test_data(test_name, file_type)
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        target_sheet_name = "Target"
        info_item = TemplateInfoItem(
            target_sheet_name, Cell(2, 2), "TBL_1",
            ["C", "B", "A"], row
        )
        sut = SUT(info_item, config)

        # when
        sut.write(data_frame, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, target_sheet_name, test_name)

    def test_write_vertical_header(self):
        # given
        test_name = "WriteWithVerticalHeader"
        file_type = "excel"
        data_frame = get_test_data(test_name, file_type)
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        target_sheet_name = "Target"
        info_item = TemplateInfoItem(
            target_sheet_name, Cell(2, 2), "TBL_1",
            ["C", "B", "A"], column
        )
        sut = SUT(info_item, config)

        # when
        sut.write(data_frame, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, target_sheet_name, test_name)

    def test_write_without_header(self):
        # given
        test_name = "WriteWithoutHeader"
        file_type = "excel"
        data_frame = get_test_data(test_name, file_type)
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        target_sheet_name = "Target"
        info_item = TemplateInfoItem(
            target_sheet_name, Cell(2, 2), "TBL_1",
            ["C", "B", "A"], row
        )
        sut = SUT(info_item, config)

        # when
        sut.write(data_frame, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, target_sheet_name, test_name)

    def test_write_limited_rows(self):
        # given
        test_name = "WriteWithLimitedRow"
        file_type = "excel"
        data_frame = get_test_data(test_name, file_type)
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        target_sheet_name = "Target"
        info_item = TemplateInfoItem(
            target_sheet_name, Cell(2, 2), "TBL_1",
            ["C", "B", "A"], row
        )
        sut = SUT(info_item, config)

        # when
        sut.write(data_frame, self.__test_target_path)

        # then
        self.assert_expected_equal_to_result(file_type, target_sheet_name, test_name)

    def test_will_raise_exception_when_data_not_have_the_needed_column(self):
        # given
        test_name = "WillRaiseExceptionWhenDataFrameNotProvideNeededColumn"
        file_type = "excel"
        data_frame = get_test_data(test_name, file_type)
        config = read_writer_config(test_name, file_type)
        self.__test_target_path = get_test_target_file_path(test_name, file_type)
        target_sheet_name = "Target"
        info_item = TemplateInfoItem(
            target_sheet_name, Cell(2, 2), "TBL_1",
            ["C", "B", "A"], row
        )
        sut = SUT(info_item, config)

        # when
        def func():
            sut.write(data_frame, self.__test_target_path)

        self.assertRaises(Exception, func)

    def assert_expected_equal_to_result(self, file_type, target_sheet_name, test_name):
        self.assertTrue(
            path.exists(self.__test_target_path)
        )
        assert_data_in_excel_is_equal(
            self,
            get_expected_target_file_path(test_name, file_type),
            self.__test_target_path,
            target_sheet_name
        )


def get_test_data_prefix(test_name, file_type):
    return path.join(
        path.dirname(__file__), "test_data", "ExcelDataFrameWriter", file_type, test_name
    )


def get_test_data(test_name, file_type):
    file_path = path.join(
        get_test_data_prefix(test_name, file_type), "Data.xlsx"
    )
    return pd.read_excel(
        file_path
    )


def get_test_target_file_path(test_name, file_type):
    suffix = "xlsx" if file_type == "excel" else "csv"
    return path.join(
        get_test_data_prefix(test_name, file_type),
        "{}.{}".format("TmpOut", suffix)
    )


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
