import openpyxl
import pandas as pd
from pandas import DataFrame

from xl_transform.common import TemplateInfoItem, row
from xl_transform.common.ConfigPaseUtils import get_cell_value_convert_func
from xl_transform.reader import ExcelDataExtractor


class DataFrameReader(object):

    def __init__(self, info_item, config):
        """

        :param TemplateInfoItem info_item:
        :param dict[str,object] config:
        """
        self.__sheet_name = info_item.sheet_name
        self.__top_left_point = info_item.top_left_point
        self.__mapping_name = info_item.mapping_name
        self.__headers = list(info_item.headers)
        self.__header_direction = info_item.header_direction
        self.__skip_header = True if "skip_header" in config and "true" == config["skip_header"].lower() else False
        self.__rows_limit = int(config["rows_limit"]) if "rows_limit" in config else None
        self.__type_hints = parse_type_hint_config(config)

    def read(self, source_file_path):
        """

        :param str source_file_path:
        :return: mapping name and data frame
        :rtype: (str,DataFrame)
        """
        if self.__sheet_name:
            # read from excel
            wb = openpyxl.load_workbook(source_file_path)
            # read header from source when "_" exists in header list
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

        else:
            df = pd.read_csv(
                source_file_path
                , header=None, dtype=str
            )

        return self.__proeccd_df(df)

    def __proeccd_df(self, df):
        unbounded_data = self.__get_unbounded_data(df)
        # support the feature: can auto detect the header, when input such symbol "${<mapping_name>:_}".
        headers = [
            header_in_info_item if "_" != header_in_info_item else header_in_data
            for header_in_data, header_in_info_item
            in zip(unbounded_data.iloc[0].values, self.__headers)
        ]
        self.__headers = headers
        columns_rename_map = {
            old: new for old, new in zip(unbounded_data.columns.values, headers)
        }
        unbounded_data.rename(columns=columns_rename_map, inplace=True)
        if "skip_header" in self.__config and "true" == self.__config["skip_header"].lower():
            unbounded_data = unbounded_data[1:]
        bounded_data = self.__get_bounded_data(unbounded_data)
        self.__update_cells_data_type(bounded_data)
        bounded_data.reindex(index=range(0, unbounded_data.shape[0]))
        return self.__mapping_name, bounded_data

    def __get_bounded_data(self, unbounded_data):
        if "rows_limit" in self.__config:
            max_ind = int(self.__config["rows_limit"])
            unbounded_data = unbounded_data[:max_ind]

        def get_last_not_null_cell_index(series):
            for ind in range(len(series) - 1, 0, -1):
                if not series.iloc[ind]:
                    return ind
            return 0

        last_not_null_cell_index_series = unbounded_data.isnull().apply(
            get_last_not_null_cell_index, axis=0
        )
        max_not_null_cell_index = sorted(last_not_null_cell_index_series.values, reverse=True)[0]
        bounded_data = unbounded_data.iloc[0:max_not_null_cell_index + 1]
        return bounded_data

    def __get_unbounded_data(self, df):
        """

        :param DataFrame df:
        :return:
        :rtype: DataFrame
        """
        x_start, y_start = self.__top_left_point.x, self.__top_left_point.y
        if row == self.__header_direction:
            x_end = df.shape[0]
            y_end = self.__top_left_point.y + len(self.__headers)
            return df.iloc[x_start:x_end, y_start:y_end]
        else:
            x_end = self.__top_left_point.x + len(self.__headers)
            y_end = df.shape[1]
            return df.iloc[x_start:x_end, y_start:y_end].T

    def __update_cells_data_type(self, df):
        """

        :param DataFrame df:
        :return:
        """
        if "type_hints" not in self.__config:
            return
        type_hints = self.__config["type_hints"]

        max_row_idx = df.shape[0]
        for header_name, type_hint in type_hints.items():
            if header_name not in self.__headers:
                continue
            column_idx = self.__headers.index(header_name)
            transfer_func = get_cell_value_convert_func(type_hint)
            if transfer_func is None:
                continue
            for row_idx in range(0, max_row_idx):
                old_val = df.iloc[row_idx].iloc[column_idx]
                new_value = transfer_func(old_val) if old_val is not None else old_val
                df.iloc[row_idx].iloc[column_idx] = new_value


def parse_type_hint_config(config):
    if "type_hints" in config:
        type_hints_dict = {header: get_cell_value_convert_func(hint_str)
                           for header, hint_str in config["type_hints"].items()}
        return {k: v for k, v in type_hints_dict.items() if v is not None}
    return None
