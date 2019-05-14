import openpyxl

from xl_transform.common import Cell, Template


class CellValueExtractor(object):

    def __init__(self, template, config):
        """

        :param Template template:
        :param dict[str,object] config:
        """
        self.__info_dict = {}  # type: dict[str,list[CellExtractInfoItem]]

        cell_info_items = template.cell_info_items

        for cell_info_item in cell_info_items:
            sheet_name = cell_info_item.sheet_name
            param_name = cell_info_item.param_name
            cell = Cell(
                cell_info_item.cell.x,
                cell_info_item.cell.y
            )
            type_convert = None
            if config is not None and 'type_hints' in config:
                type_hints = config['type_hints']
                if param_name in type_hints:
                    from xl_transform.common.ConfigPaseUtils import get_cell_value_convert_func
                    type_convert = get_cell_value_convert_func(type_hints[param_name])

            cell_extract_info_item = CellExtractInfoItem(
                param_name,
                cell,
                type_convert
            )

            info_items_in_sheet = self.__info_dict.setdefault(sheet_name, [])
            info_items_in_sheet.append(cell_extract_info_item)

    def read(self, source_excel):
        """

        :param str or WorkBook source_excel:
        :return:
        :rtype: dict[str,object]
        """
        if isinstance(source_excel, str):
            wb = openpyxl.load_workbook(source_excel)
        else:
            wb = source_excel

        result = {}

        for sheet in wb:
            sheet_name = sheet.title
            if sheet_name in self.__info_dict:
                info_items_in_sheet = self.__info_dict[sheet_name]
                for info_item in info_items_in_sheet:  # type: CellExtractInfoItem
                    cell = sheet.cell(
                        info_item.cell.x, info_item.cell.y
                    )
                    if info_item.type_convert is not None:
                        result[info_item.param_name] = info_item.type_convert(cell)
                    else:
                        result[info_item.param_name] = cell.value

        return result


class CellExtractInfoItem(object):

    def __init__(
            self,
            param_name,
            cell,
            type_convert
    ):
        """

        :param str param_name:
        :param Cell cell:
        :param (xlCell)->Any type_convert:
        """
        self.param_name = param_name
        self.cell = cell
        self.type_convert = type_convert
