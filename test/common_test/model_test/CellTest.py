import unittest

from common.model import Cell


class CellTest(unittest.TestCase):

    def test_return_true_when_contact_in_the_same_row(self):
        # given
        a, b = Cell(1, 1), Cell(1, 2)

        # then
        self.assert_mutually_contact(a, b)

    def test_return_true_when_contact_in_the_same_column(self):
        # given
        a, b = Cell(1, 1), Cell(2, 1)

        # then
        self.assert_mutually_contact(a, b)

    def test_return_false_when_contact_in_scenario_1(self):
        """
        Return false when the position like below grape.

          A#
          #B

        :return:
        """
        # given
        a, b = Cell(1, 1), Cell(2, 2)

        # then
        self.assert_not_mutually_contact(a, b)

    def test_equal_when_x_and_y_is_equal(self):
        # given
        a, b = Cell(1, 1), Cell(1, 1)

        # then
        self.assertEqual(a, b)

    def test_equal_when_x_and_y_is_equal_by_hash(self):
        # given
        a, b = Cell(1, 1), Cell(1, 1)

        # then
        self.assertEqual(1, len({a, b}))

    def assert_mutually_contact(self, a, b):
        # type Cell a
        result1, result2 = a.is_contact_with(b), b.is_contact_with(a)

        self.assertTrue(result1)
        self.assertTrue(result2)

    def assert_not_mutually_contact(self, a, b):
        # type Cell a
        result1, result2 = a.is_contact_with(b), b.is_contact_with(a)

        self.assertFalse(result1)
        self.assertFalse(result2)
