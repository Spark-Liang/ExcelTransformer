import re

import pandas as pd

from common import CellsSetUtils
from common.model import Cell
from common.model import OrthorhombicLine

global row, column

row = "row"
column = "column"


class Template(object):

    def __init__(self, template_file_path, type_hint=None):
        """

        :param str template_file_path:
        :param str type_hint: directly to restrict the file type,
            otherwise the program will detect the type by filename extension. \n
            Currently support "excel" or "csv".
        """
        pass


class TemplateInfoItem(object):
    """
    Each one of the "TemplateInfoItem" represent a single template area. This class contain the below properties to describe the template information.
    Properties
    __________
    sheet_name : str
        The sheet where the template area is located in. \n
        This value will be available only when the template in an Excel file.

    top_left_point : Cell
        This property is used to location the template area.

    mapping_name : str
        The name of the associated mapping.

    headers : list[str]
        The list of the column headers exists in the template area.

    header_direction : str
        The direction of the header. It can be "row" or "column".


    """

    def __init__(
            self,
            sheet_name, top_left_point, mapping_name, headers, header_direction=row
    ):
        """

        :param str sheet_name:
        :param Cell top_left_point:
        :param str mapping_name:
        :param list[str] headers:
        :param str header_direction: "row" or "column".
        """
        self.__sheet_name = sheet_name
        self.__top_left_point = top_left_point
        self.__mapping_name = mapping_name
        self.__headers = tuple(headers)
        self.__header_direction = header_direction

    @property
    def top_left_point(self):
        return self.__top_left_point

    @property
    def sheet_name(self):
        return self.__sheet_name

    @property
    def mapping_name(self):
        return self.__mapping_name

    @property
    def headers(self):
        return self.__headers

    @property
    def header_direction(self):
        return self.__header_direction

    def __eq__(self, other):
        return (self.top_left_point == other.top_left_point
                and self.sheet_name == other.sheet_name
                and self.mapping_name == other.mapping_name
                and self.headers == other.headers
                and self.header_direction == other.header_direction)

    def __hash__(self):
        return hash((
            self.top_left_point,
            self.sheet_name,
            self.mapping_name,
            self.headers,
            self.header_direction
        ))


def parse_excel_template(filename):
    """

    :param str filename:
    :return:
    :rtype: list[TemplateInfoItem]
    """
    result = []

    df_dict = pd.read_excel(filename, sheet_name=None, header=None)
    for sheet_name, df in df_dict.items():
        result.extend(
            __parse_data_frame(df, sheet_name)
        )
    return result


def parse_csv_template(filename):
    """

    :param str filename:
    :return:
    :rtype: list[TemplateInfoItem]
    """
    return __parse_data_frame(
        pd.read_csv(filename, header=None),
        None
    )


def __parse_data_frame(df, sheet_name):
    max_x, max_y = df.shape
    target_cell_info_dict = {}
    for x in range(0, max_x):
        for y in range(0, max_y):
            cell_value = df.at[x, y]
            if isinstance(cell_value, str):
                __parse_cell_info_into_dict(
                    x, y, cell_value, target_cell_info_dict
                )
    target_cells = target_cell_info_dict.keys()
    list_of_contact_cells_set = \
        CellsSetUtils.separate_cells_set_into_contacted_cells_set(target_cells)
    # check if all the shape of the contacted cells are orthorhombic line.
    return __check_line_shape_and_construct_parse_result_items(
        sheet_name, list_of_contact_cells_set, target_cell_info_dict
    )


def __parse_cell_info_into_dict(x, y, value, target_cell_info_dict):
    pattern = r"^\s*\$\{([^}]+?)\:([^}]+?)\}\s*$"
    matchObj = re.match(pattern, value)
    if matchObj is None:
        return
    mapping_name = matchObj.group(1)
    header_name = matchObj.group(2)
    target_cell_info_dict[Cell(x, y)] = {
        "mapping_name": mapping_name,
        "header_name": header_name
    }


def __check_line_shape_and_construct_parse_result_items(
        sheet_name, contact_cells_set_list, target_cell_info_dict
):
    """

    :param list[set[Cell]] contact_cells_set_list:
    :param dict[Cell,dict[str,str]] target_cell_info_dict:
    :return:
    :rtype: list[TemplateInfoItem]
    """
    result = []
    for contact_cells_set in contact_cells_set_list:
        mapping_cells_dict = {}
        for cell in contact_cells_set:
            mapping_name = target_cell_info_dict[cell]["mapping_name"]
            mapping_cell_set = mapping_cells_dict.setdefault(
                mapping_name, set()
            )
            mapping_cell_set.add(cell)
        for mapping, cells_set_of_mapping in mapping_cells_dict.items():
            line = OrthorhombicLine(cells_set_of_mapping)
            result.append(TemplateInfoItem(
                sheet_name,
                line.cells_list[0],
                mapping,
                [target_cell_info_dict[cell]["header_name"] for cell in line.cells_list],
                column if line.is_vertical else row
            ))
    return result
