import unittest

from xl_transform.common.model import CellsArea


class CellsAreaTest(unittest.TestCase):

    def test_return_false_when_two_area_not_intersect(self):
        # given
        a, b = CellsArea((1, 1), (3, 4), "A"), CellsArea((4, 5), (5, 6), "A")

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
        a, b = CellsArea((1, 1), (3, 4), "A"), CellsArea((4, 1), (5, 4), "A")

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
        a, b = CellsArea((1, 1), (3, 4), "A"), CellsArea((1, 5), (3, 8), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertFalse(result1)
        self.assertFalse(result2)

    def test_return_false_when_two_area_not_in_same_sheet(self):
        # given
        a, b = CellsArea((1, 1), (3, 4), "A"), CellsArea((1, 1), (3, 4), "B")

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
        a, b = CellsArea((1, 1), (4, 4), "A"), CellsArea((3, 3), (6, 6), "A")

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
        a, b = CellsArea((4, 1), (6, 6), "A"), CellsArea((1, 3), (6, 5), "A")

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
        a, b = CellsArea((1, 1), (6, 6), "A"), CellsArea((3, 3), (5, 5), "A")

        # when
        result1, result2 = a.is_intersect_with(b), b.is_intersect_with(a)

        # then
        self.assertTrue(result1)
        self.assertTrue(result2)
