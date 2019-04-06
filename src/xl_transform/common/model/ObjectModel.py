import math


class CellsArea(object):

    def __init__(self, top_left_point, bottom_right_point, sheet_name):
        """

        :param (int,int) or Cell top_left_point: the position of the top left of the area.
        :param (int,int) or Cell bottom_right_point:
        :param str sheet_name:
        """
        if (top_left_point[0] < 0 or top_left_point[1] < 0
                or bottom_right_point[0] < 0 or bottom_right_point[1] < 0):
            raise Exception("The x index and the y index must greater than or equal to 0")
        if (top_left_point[0] >= bottom_right_point[0]
                or top_left_point[1] >= bottom_right_point[1]):
            raise Exception("The top_left_point most at the top left side of the bottom_right_point!")
        self.top_left_point = Cell(top_left_point[0], top_left_point[1])
        self.bottom_right_point = Cell(bottom_right_point[0], bottom_right_point[1])
        self.sheet_name = sheet_name

    def is_intersect_with(self, other):
        """
        Check this CellsArea is contect with other CellsArea.
        :param CellsArea other:
        :return:
        :rtype: bool
        """
        if self.sheet_name != other.sheet_name:
            return False
        self_center, other_center = self.__get_center(), other.__get_center()
        dx = math.fabs(self_center[0] - other_center[0])
        dy = math.fabs(self_center[1] - other_center[1])
        max_dx = math.fabs((self.__get_width() + other.__get_width()) / 2)
        max_dy = math.fabs((self.__get_height() + other.__get_height()) / 2)
        return dx < max_dx and dy < max_dy

    def __get_center(self):
        """

        :return:
        """
        x = (self.top_left_point[0] + self.bottom_right_point[0]) / 2
        y = (self.top_left_point[1] + self.bottom_right_point[1]) / 2
        return x, y

    def __get_width(self):
        return self.bottom_right_point[0] - self.top_left_point[0]

    def __get_height(self):
        return self.bottom_right_point[1] - self.top_left_point[1]


class Cell(object):

    def __init__(self, x, y):
        """

        :param int x: Row index of the cell. The index starts with 1.
        :param int y: Column index of the cell.The index starts with 1.
        """
        if x < 1:
            err_msg = "The row index of cell must greater than or equal to 1. Current is {}".format(x)
            raise Exception(err_msg)
        if y < 1:
            err_msg = "The column index of cell must greater than or equal to 1. Current is {}".format(y)
            raise Exception(err_msg)
        self.__x = x
        self.__y = y

    @property
    def x(self):
        return self.__x

    @property
    def y(self):
        return self.__y

    def is_contact_with(self, other):
        # contact when in same row
        if self.x == other.x and 1 == math.fabs(self.y - other.y):
            return True
        # contact when in same column
        elif self.y == other.y and 1 == math.fabs(self.x - other.x):
            return True
        else:
            return False

    def move(self, direction, distance=1, reverse=False):
        """

        :param str direction:
        :param int distance:
        :param bool reverse:
        :return:
        """
        if direction == "row":
            return Cell(self.__x, self.__y + distance * (1 if not reverse else -1))
        elif direction == "column":
            return Cell(self.__x + distance * (1 if not reverse else -1), self.__y)

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False
        return self.x == other.x and self.y == other.y

    def __hash__(self):
        return hash((self.__x, self.__y))

    def __str__(self):
        return "Cell({},{})".format(self.x, self.y)

    def __getitem__(self, item):
        if item == 0:
            return self.__x
        elif item == 1:
            return self.__y
        else:
            err_msg = "Only support '0' and '1' as the index"
            raise Exception(err_msg)


class OrthorhombicLine(object):
    """
    A object model to describe the shape of a certain cell set.

    Properties
    __________
    is_vertical : bool or None
        is used to describe the direction of the OrthorhombicLine. \n
        When OrthorhombicLine only contain single cell or zero cell, it will be None.

    cells_list : list[Cell]
        the list of the cell in this OrthorhombicLine. \n
        The order of the cell will be from left to right and from top to bottom.
    """

    def __init__(self, cells):
        """

        :param Iterable[Cell] cells:
        """
        if len(cells) == 0:
            self.__is_vertical = None
            self.__cells_list = []
        elif len(cells) == 1:
            self.__is_vertical = None
            self.__cells_list = list(cells)
        else:
            self.__is_vertical = OrthorhombicLine.__parse_direction(cells)
            if self.is_vertical:
                sort_key = lambda cell: cell.x
            else:
                sort_key = lambda cell: cell.y
            self.__cells_list = sorted(
                cells, key=sort_key
            )

    @staticmethod
    def __parse_direction(cells):
        """

        :param Iterable[Cell] cells:
        :return:
        :rtype: bool
        """
        err_msg = "The given cells can not construct OrthorhombicLine."

        cells_list = list(cells)
        if (cells_list[0].x != cells_list[1].x
                and cells_list[0].y != cells_list[1].y):
            # not vertical or horizontal
            raise Exception(err_msg)
        try_vertical = cells_list[0].y == cells_list[1].y
        if try_vertical:
            y = cells_list[0].y
            for cell in cells:
                if cell.y != y:
                    raise Exception(err_msg)
        else:
            x = cells_list[0].x
            for cell in cells:
                if cell.x != x:
                    raise Exception(err_msg)

        return try_vertical

    @property
    def is_vertical(self):
        return self.__is_vertical

    @property
    def cells_list(self):
        return list(self.__cells_list)
