import unittest
from datetime import datetime

from xl_transform.common import Template
from xl_transform.reader import CellValueExtractor as SUT


class CellValueExtractorTest(unittest.TestCase):

    def test_can_read_single_cell_with_type_hints(self):
        # given
        test_name = 'CanReadSingleCellWithTypehints'
        source_data_path = get_test_source_data_path(test_name)
        sut = SUT(
            Template(get_template_path(test_name)),
            get_config_data(test_name)
        )

        # when
        result = sut.read(source_data_path)

        # then
        self.assertDictEqual(
            {
                "test_param1": datetime.strptime('2019/12/21', r'%Y/%m/%d'),
                "test_param2": datetime.strptime('2019/12/23', r'%Y/%m/%d')
            },
            result
        )


from os import path


def get_test_data_prefix(test_name):
    return path.join(
        path.dirname(__file__), "test_data", "CellValueExtractor", test_name
    )


def get_test_source_data_path(test_name):
    return path.join(
        get_test_data_prefix(test_name),
        "Source.xlsx"
    )


def get_template_path(test_name):
    return path.join(
        get_test_data_prefix(test_name),
        "Template.xlsx"
    )


def get_config_data(test_name):
    import json
    with open(
            path.join(
                get_test_data_prefix(test_name),
                "config.json"
            )
    )as config_file:
        return json.load(config_file)
