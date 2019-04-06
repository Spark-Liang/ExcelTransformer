import unittest

from xl_transform.common import Cell
from xl_transform.common import CellsSetUtils as SUT


class Test_method_separate_cells_set_into_cells_area(unittest.TestCase):

    def test_can_separate_cells_set_into_cells_area(self):
        """
        The cells location showed in below graph.

          1 2 3 4 5 6 7 8
        1######
        2         #####
        3   #     #
        4   #     #
        5   #
        6      ######
        7      ######
        8      ######

        :return:
        """
        # given

        set_for_horizontal_line = {Cell(x, y) for x, y in [(1, 1), (1, 2), (1, 3)]}
        set_for_vertical_line = {Cell(x, y) for x, y in [(3, 2), (4, 2), (5, 2)]}
        set_for_L_shape_line = {Cell(x, y) for x, y in [(4, 5), (3, 5), (2, 5), (2, 6), (2, 7)]}
        set_for_rectangle_shape_area = {Cell(x, y) for x in range(6, 9) for y in range(4, 7)}
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

          1 2 3 4 5 6 7 8
        1######
        2     ### 1
        3 #
        4 #
        5 ### 2
        6   #
        7   #
        8
        :return:
        """
        # given

        set_for_scenario_1 = {Cell(x, y) for x, y in [(1, 1), (1, 2), (1, 3), (2, 3), (2, 4)]}
        set_for_scenario_2 = {Cell(x, y) for x, y in [(3, 1), (4, 1), (5, 1), (5, 2), (6, 2)]}
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
