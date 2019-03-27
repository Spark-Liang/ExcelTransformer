import unittest

from common.model import Cell
from common.model import OrthorhombicLine as SUT


class OrthorhombicLineTest(unittest.TestCase):

    def test_when_create_with_zero_cell(self):
        # when
        result = SUT(set())

        # then
        self.assertEqual(None, result.is_vertical)
        self.assertEqual([], result.cells_list)

    def test_when_create_with_single_cell(self):
        # given
        single_cell = Cell(1, 1)

        # when
        result = SUT({single_cell})

        # then
        self.assertEqual(None, result.is_vertical)
        self.assertEqual([single_cell], result.cells_list)

    def test_when_create_with_two_cell_in_horizontal_direction(self):
        # given
        cell1, cell2 = Cell(1, 1), Cell(1, 2)

        # when
        result = SUT({cell1, cell2})

        # then
        self.assertEqual(False, result.is_vertical)
        self.assertEqual([cell1, cell2], result.cells_list)

    def test_when_create_with_three_cell_in_horizontal_direction(self):
        # given
        cell1, cell2, cell3 = Cell(1, 1), Cell(1, 2), Cell(1, 3)

        # when
        result = SUT({cell1, cell2, cell3})

        # then
        self.assertEqual(False, result.is_vertical)
        self.assertEqual([cell1, cell2, cell3], result.cells_list)

    def test_when_create_with_two_cell_in_vertical_direction(self):
        # given
        cell1, cell2 = Cell(1, 1), Cell(2, 1)

        # when
        result = SUT({cell1, cell2})

        # then
        self.assertEqual(True, result.is_vertical)
        self.assertEqual([cell1, cell2], result.cells_list)

    def test_when_create_with_three_cell_in_vertical_direction(self):
        # given
        cell1, cell2, cell3 = Cell(1, 1), Cell(2, 1), Cell(3, 1)

        # when
        result = SUT({cell1, cell2, cell3})

        # then
        self.assertEqual(True, result.is_vertical)
        self.assertEqual([cell1, cell2, cell3], result.cells_list)

    def test_raise_exception_when_is_not_a_shape_of_orthorhombic_1(self):
        """
        The shape of the cells is like the below graph.

          0 1 2 3
        0 #
        1   #
        2
        3

        :return:
        """
        # given
        cell1, cell2 = Cell(1, 1), Cell(2, 2)

        # when

        def sut():
            SUT({cell1, cell2})

        self.assertRaises(Exception, sut)

    def test_raise_exception_when_is_not_a_shape_of_orthorhombic_2(self):
        """
        The shape of the cells is like the below graph.

          0 1 2 3
        0 #
        1 ###
        2
        3

        :return:
        """

        def sut():
            SUT({Cell(1, 1), Cell(2, 1), Cell(2, 2)})

        self.assertRaises(Exception, sut)
