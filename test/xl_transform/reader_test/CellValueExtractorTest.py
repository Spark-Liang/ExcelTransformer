import unittest
from datetime import datetime

from xl_transform.reader import CellValueExtractor as SUT


class CellValueExtractorTest(unittest.TestCase):

    def test_can_read_single_cell_with_type_hints(self):
        # given
        test_name = 'CanReadSingleCellWithTypehints'
        source_data_path = get_test_source_data_path(test_name)
        sut = SUT()

        # when
        result = sut.read(source_data_path)

        # then
        result_value = result['test_param']
        self.assertEqual(
            result_value,
            datetime.strptime(r'%Y/%m/%d', '2019/12/21')
        )


from os import path


def get_test_data_prefix(test_name):
    return path.join(
        path.dirname(__file__), "test_data", "CellValueExtractor", test_name
    )


def get_test_source_data_path(testname):
    return path.join(
        get_test_data_prefix(testname),
        "Source.xlsx"
    )
