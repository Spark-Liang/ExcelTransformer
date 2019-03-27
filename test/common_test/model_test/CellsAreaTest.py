import unittest

from common.model import CellsArea


class CellsAreaTest(unittest.TestCase):

    def test_return_false_when_two_area_not_intersect(self):
        # given
        a, b = CellsArea((0, 0), (2, 3), "A"), CellsArea((3, 4), (4, 5), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertFalse(result1)
        self.assertFalse(result2)

    def test_return_false_when_two_area_just_horizontal_contact_but_not_intersect(self):
        """
        Scenario is showed in the below picture.
           ##########
           #   ##   #
           # a ## b #
           #   ##   #
           ##########

        """
        # given
        a, b = CellsArea((0, 0), (2, 3), "A"), CellsArea((3, 0), (4, 3), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertFalse(result1)
        self.assertFalse(result2)

    def test_return_false_when_two_area_just_vertical_contact_but_not_intersect(self):
        """
        Scenario is showed in the below picture.
           #####
           #   #
           # a #
           #   #
           #####
           #####
           #   #
           # b #
           #   #
           #####
        """
        # given
        a, b = CellsArea((0, 0), (2, 3), "A"), CellsArea((0, 4), (2, 7), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertFalse(result1)
        self.assertFalse(result2)

    def test_return_false_when_two_area_not_in_same_sheet(self):
        # given
        a, b = CellsArea((0, 0), (2, 3), "A"), CellsArea((0, 0), (2, 3), "B")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertFalse(result1)
        self.assertFalse(result2)

    def test_return_true_with_intersect_scenario_1(self):
        """
        Scenario is showed in the below picture.
           #####
           #   #
           # a #
           # #####
           ##### #
             # b #
             #   #
             #####

        """
        # given
        a, b = CellsArea((0, 0), (3, 3), "A"), CellsArea((2, 2), (5, 5), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertTrue(result1)
        self.assertTrue(result2)

    def test_return_true_with_intersect_scenario_2(self):
        """
        Scenario is showed in the below picture.
               #####
               #   #
               # a #
            ############
            #  #   #   #
            #  #   # b #
            ############
               #####

        """
        # given
        a, b = CellsArea((3, 0), (5, 5), "A"), CellsArea((0, 2), (5, 4), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertTrue(result1)
        self.assertTrue(result2)

    def test_return_true_with_intersect_scenario_3(self):
        """
        Scenario is showed in the below picture.
          ###########
          #         #
          # a       #
          #  ###### #
          #  #    # #
          #  #  b # #
          #  ###### #
          #         #
          ###########


        """
        # given
        a, b = CellsArea((0, 0), (5, 5), "A"), CellsArea((2, 2), (4, 4), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertTrue(result1)
        self.assertTrue(result2)
