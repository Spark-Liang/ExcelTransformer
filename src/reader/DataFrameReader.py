import pandas as pd
from pandas import DataFrame

from common import TemplateInfoItem, row


class DataFrameReader(object):

    def __init__(self, info_item, config):
        """

        :param TemplateInfoItem info_item:
        :param dict[str,object] config:
        """
        self.__sheet_name = info_item.sheet_name
        self.__top_left_point = info_item.top_left_point
        self.__mapping_name = info_item.mapping_name
        self.__headers = info_item.headers
        self.__header_direction = info_item.header_direction
        self.__config = config

    def read(self, source_file_path):
        """

        :param str source_file_path:
        :return:
        :rtype: (str,DataFrame)
        """
        if self.__sheet_name:
            df = pd.read_excel(
                source_file_path,
                sheet_name=self.__sheet_name,
                header=None,
                na_values=None
            )
        else:
            df = pd.read_csv(source_file_path, header=None)

        unbounded_data = self.__get_unbounded_data(df)

        headers = [
            header_in_info_item if "_" != header_in_info_item else header_in_data
            for header_in_data, header_in_info_item
            in zip(unbounded_data.iloc[0].values, self.__headers)
        ]

        columns_rename_map = {
            old: new for old, new in zip(unbounded_data.columns.values, headers)
        }
        unbounded_data.rename(columns=columns_rename_map, inplace=True)

        if "skip_header" in self.__config and bool(self.__config["skip_header"]):
            unbounded_data = unbounded_data[1:]

        if "rows_limit" in self.__config:
            max_ind = int(self.__config["rows_limit"])
            unbounded_data = unbounded_data[:max_ind]

        bounded_data = self.__get_bounded_data(unbounded_data)

        bounded_data.reindex(index=range(0, unbounded_data.shape[0]))
        return self.__mapping_name, bounded_data

    @staticmethod
    def __get_bounded_data(unbounded_data):
        def get_last_not_null_cell_index(series):
            for ind in range(len(series) - 1, 0, -1):
                if not series.iloc[ind]:
                    return ind

        last_not_null_cell_index_series = unbounded_data.isnull().apply(
            get_last_not_null_cell_index, axis=0
        )
        max_not_null_cell_index = sorted(last_not_null_cell_index_series.values, reverse=True)[0]
        bounded_data = unbounded_data[0:max_not_null_cell_index + 1]
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
