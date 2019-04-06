import openpyxl
from pandas import DataFrame

from xl_transform.common import TemplateInfoItem, row, column
from xl_transform.common.ConfigPaseUtils import get_cell_value_convert_func
from xl_transform.reader import ExcelDataExtractor


class DataFrameReader(object):

    def __init__(self, info_item, config=None):
        """

        :param TemplateInfoItem info_item:
        :param dict[str,object] config:
        """
        self.__sheet_name = info_item.sheet_name
        self.__top_left_point = info_item.top_left_point
        self.__mapping_name = info_item.mapping_name
        self.__headers = list(info_item.headers)
        if info_item.header_direction is not None:
            self.__header_direction = info_item.header_direction
        else:
            if "data_column_direction" not in config:
                err_msg = """Please config the 'data_column_direction' when extract only one data column. 
                          The problematic mapping is {}""".format(self.__mapping_name)
                raise Exception(err_msg)
            data_column_direction = config["data_column_direction"].lower()
            self.__header_direction = row if column == data_column_direction else column
        # The default value of "skip_header" is True
        self.__skip_header = False if config is not None and "skip_header" in config and "false" == config[
            "skip_header"].lower() else True
        self.__rows_limit = int(config["rows_limit"]) if config is not None and "rows_limit" in config else None
        self.__type_hints = parse_type_hint_config(config) if config is not None else None

    def read(self, source_file_path):
        """

        :param str source_file_path:
        :return: mapping name and data frame
        :rtype: (str,DataFrame)
        """
        if self.__sheet_name:
            # read_data from excel
            wb = openpyxl.load_workbook(source_file_path)
            # read_data header from source when "_" exists in header list
            if "_" in self.__headers:
                headers = ExcelDataExtractor.extract_header(
                    wb[self.__sheet_name],
                    start_row_idx=self.__top_left_point.x + 1,
                    start_col_idx=self.__top_left_point.y + 1,
                    header_row_direction=self.__header_direction,
                    data_column_limit=len(self.__headers)
                )
            else:
                headers = self.__headers

            if self.__skip_header:
                if self.__header_direction == row:
                    start_row_idx = self.__top_left_point.x + 2
                    start_col_idx = self.__top_left_point.y + 1
                else:
                    start_row_idx = self.__top_left_point.x + 1
                    start_col_idx = self.__top_left_point.y + 2
            else:
                start_row_idx = self.__top_left_point.x + 1
                start_col_idx = self.__top_left_point.y + 1

            return self.__mapping_name, ExcelDataExtractor.extract_data_frame(
                wb, self.__sheet_name,
                header_list=headers,
                start_row_idx=start_row_idx,
                start_col_idx=start_col_idx,
                import_header=False,
                data_row_direction=self.__header_direction,
                data_rows_limit=self.__rows_limit,
                converter=self.__type_hints
            )


def parse_type_hint_config(config):
    if "type_hints" in config:
        type_hints_dict = {header: get_cell_value_convert_func(hint_str)
                           for header, hint_str in config["type_hints"].items()}
        return {k: v for k, v in type_hints_dict.items() if v is not None}
    return None
