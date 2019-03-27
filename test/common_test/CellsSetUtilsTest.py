import unittest

from common import CellsSetUtils as SUT
from common.model import Cell


class Test_method_separate_cells_set_into_cells_area(unittest.TestCase):

    def test_can_separate_cells_set_into_cells_area(self):
        """
        The cells location showed in below graph.

          0 1 2 3 4 5 6 7
        0######
        1         #####
        2   #     #
        3   #     #
        4   #
        5      ######
        6      ######
        7      ######

        :return:
        """
        # given

        set_for_horizontal_line = {Cell(x, y) for x, y in [(0, 0), (0, 1), (0, 2)]}
        set_for_vertical_line = {Cell(x, y) for x, y in [(2, 1), (3, 1), (4, 1)]}
        set_for_L_shape_line = {Cell(x, y) for x, y in [(3, 4), (2, 4), (1, 4), (1, 5), (1, 6)]}
        set_for_rectangle_shape_area = {Cell(x, y) for x in range(5, 8) for y in range(3, 6)}
        expected_cells_set_list = [set_for_horizontal_line, set_for_vertical_line, set_for_L_shape_line,
                                   set_for_rectangle_shape_area]
        cells_set = set()
        for cells_set_tmp in expected_cells_set_list:
            cells_set |= cells_set_tmp

        # when
        result = SUT.separate_cells_set_into_contacted_cells_set(cells_set)

        # then
        for cells_set_tmp in expected_cells_set_list:
            self.assert_contain_and_remove(result, cells_set_tmp)
        self.assertEqual(len(result), 0)

    def test_can_separate_cells_set_into_cells_area_with_complex_scenario_1(self):
        """
        The cells location showed in below graph.

          0 1 2 3 4 5 6 7
        0######
        1     ### 1
        2 #
        3 #
        4 ### 2
        5   #
        6   #
        7
        :return:
        """
        # given

        set_for_scenario_1 = {Cell(x, y) for x, y in [(0, 0), (0, 1), (0, 2), (1, 2), (1, 3)]}
        set_for_scenario_2 = {Cell(x, y) for x, y in [(2, 0), (3, 0), (4, 0), (4, 1), (5, 1)]}
        expected_cells_set_list = [set_for_scenario_1, set_for_scenario_2]
        cells_set = set()
        for cells_set_tmp in expected_cells_set_list:
            cells_set |= cells_set_tmp

        # when
        result = SUT.separate_cells_set_into_contacted_cells_set(cells_set)

        # then
        for cells_set_tmp in expected_cells_set_list:
            self.assert_contain_and_remove(result, cells_set_tmp)
        self.assertEqual(len(result), 0)

    def assert_contain_and_remove(self, result, expected_cell_set):
        """
        Expected the "expected_cell_set" exists and only exists once in the result.
        And than remove the "expected_cell_set" from the result.
        :param result:
        :param expected_cell_set:
        :return:
        """
        self.assertTrue(
            result.count(expected_cell_set) == 1,
            """expected_cell_set is {},
            result is {}
            """.format(
                self.format_set(expected_cell_set),
                r"{" + ",".join([self.format_set(x) for x in result]) + r"}"
            )
        )
        result.remove(expected_cell_set)

    @staticmethod
    def format_set(set_to_format):
        return r"{" + ",".join([str(x) for x in set_to_format]) + r"}"
