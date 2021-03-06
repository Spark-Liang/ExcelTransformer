import os
import unittest
from os import path

from test_util.DataFrameAssertUtil import assert_data_in_excel_is_equal
from xl_transform.main import App as SUT


class TestTransfer(unittest.TestCase):

    def test_can_transfer_from_excel_to_excel(self):
        # given
        test_name = "TestFromExcelToExcel"
        data_source_path = get_file_path(test_name, "Source.xlsx")
        data_source_template_path = get_file_path(test_name, "SourceTemplate.xlsx")
        output_template_path = get_file_path(test_name, "OutputTemplate.xlsx")
        config_path = get_file_path(test_name, "config.json")
        output_path = get_file_path(test_name, "TestOut.xlsx")
        if path.exists(output_path):
            os.remove(output_path)

        # when
        SUT.transfer(
            data_source_path,
            data_source_template_path,
            output_template_path,
            output_path,
            config_path
        )
        # then
        self.assertTrue(path.exists(output_path))
        assert_data_in_excel_is_equal(
            self,
            get_file_path(test_name, "Result.xlsx"),
            output_path
        )

    def test_pure_transfer(self):
        # given
        test_name = "TestPureTransfer"
        data_source_path = get_file_path(test_name, "Source.xlsx")
        data_source_template_path = get_file_path(test_name, "SourceTemplate.xlsx")
        output_template_path = get_file_path(test_name, "OutputTemplate.xlsx")
        output_path = get_file_path(test_name, "TestOut.xlsx")
        if path.exists(output_path):
            os.remove(output_path)

        # when
        SUT.transfer(
            data_source_path,
            data_source_template_path,
            output_template_path,
            output_path
        )
        # then
        self.assertTrue(path.exists(output_path))
        assert_data_in_excel_is_equal(
            self,
            get_file_path(test_name, "Result.xlsx"),
            output_path
        )

    def test_can_exec_by_command(self):
        test_name = "TestFromExcelToExcel"
        data_source_path = get_file_path(test_name, "Source.xlsx")
        data_source_template_path = get_file_path(test_name, "SourceTemplate.xlsx")
        output_template_path = get_file_path(test_name, "OutputTemplate.xlsx")
        config_path = get_file_path(test_name, "config.json")
        output_path = get_file_path(test_name, "TestOut.xlsx")
        command_path = os.path.join(os.path.dirname(__file__), "..", "..", "..", "bin", "xl_transform.cmd")

        command = "{} -s {} --source-template {} -t {} --target-template {} --conf {}".format(
            command_path,
            data_source_path,
            data_source_template_path,
            output_path,
            output_template_path,
            config_path
        )
        if path.exists(output_path):
            os.remove(output_path)

        # when
        os.system(command)

        # then
        self.assertTrue(path.exists(output_path))
        assert_data_in_excel_is_equal(
            self,
            get_file_path(test_name, "Result.xlsx"),
            output_path
        )


def get_based_test_path(test_name):
    return path.join(path.dirname(__file__), "test_data", test_name)


def get_file_path(test_name, file_name):
    return path.join(get_based_test_path(test_name), file_name)
