import os

import openpyxl
from pandas import DataFrame

from xl_transform.common import Template, TemplateInfoItem, row, column
from xl_transform.common.ConfigPaseUtils import get_value_convert_func
from xl_transform.common.model import CellsArea, Cell
from xl_transform.writer import ExcelDataWriter


class ExcelDataFrameWriter(object):

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
        self.__with_header = True if "with_header" in config and "true" == config["with_header"].lower() else False
        self.__rows_limit = int(config["rows_limit"]) if "rows_limit" in config else None
        self.__format = self.parse_format_config(config)

    @property
    def sheet_name(self):
        return self.__sheet_name

    @property
    def top_left_point(self):
        return self.__top_left_point

    def write(self, df, target):
        """
        Write the given DataFrame into target area.
        :param DataFrame df :
        :param str target:
        :return:
        """
        if os.path.exists(target):
            wb = openpyxl.load_workbook(target)
        else:
            wb = openpyxl.Workbook()

        # do projection
        header_provided_in_df = df.columns.values
        for header in self.__headers:
            if header not in header_provided_in_df:
                err_msg = "The given data frame not provide the column : " + header
                raise Exception(err_msg)
        df_to_write = df[self.__headers]

        ExcelDataWriter.write_data_frame_to_workbook(
            wb, self.__sheet_name, df_to_write,
            start_row_idx=self.top_left_point.x + 1,
            start_col_idx=self.top_left_point.y + 1,
            data_row_direction=self.__header_direction,
            with_header=self.__with_header,
            data_rows_limit=self.__rows_limit,
            converter=self.__format
        )
        wb.save(target)

    def get_data_frame_to_write(self, df):
        """

        :param DataFrame df:
        :return:
        :rtype: DataFrame
        """
        unprovided_columns = set(self.__headers) - set(df.columns.values)
        if 0 != len(unprovided_columns):
            err_msg = "The given data frame does not provide these columns: " + str(unprovided_columns)
            raise Exception(err_msg)
        data_to_write = df[self.__headers]
        data_to_write = self.__format_data(data_to_write)
        if "rows_limit" in self.__config:
            max_row_idx = int(self.__config["rows_limit"]) + 1
            if max_row_idx < data_to_write.shape[0]:
                data_to_write = data_to_write.iloc[:max_row_idx]

        if "with_header" in self.__config and "true" == self.__config["with_header"].lower():
            header_frame = DataFrame(
                {header: [header_val] for header, header_val in zip(self.__headers, self.__headers)},
                columns=self.__headers
            )
            data_to_write = header_frame.append(data_to_write, ignore_index=True)
        if column == self.__header_direction:
            data_to_write = data_to_write.T
        return data_to_write

    def __format_data(self, df):
        if "format" not in self.__config:
            return df
        format_config = self.__config["format"]
        format_dict = {}
        if "*" in format_config:
            format_str = format_config["*"]
            for header in self.__headers:
                format_dict[header] = format_str
        format_dict.update({
            k: v for k, v in format_config.items() if k != "*"
        })

        max_row_idx = df.shape[0]
        new_values = {}
        for column_idx in range(0, len(self.__headers)):
            header_name = self.__headers[column_idx]
            transfer_func = get_value_convert_func(format_dict[header_name]) if header_name in format_dict else None
            new_value_list = []
            for row_idx in range(0, max_row_idx):
                old_val = df.iloc[row_idx].iloc[column_idx]
                if (transfer_func is not None
                        and old_val is not None):
                    new_value_list.append(transfer_func(old_val))
                else:
                    new_value_list.append(old_val)
            new_values[header_name] = new_value_list
        return DataFrame(
            new_values, columns=self.__headers, dtype=str
        )

    def get_write_area(self, df):
        """
        Get the area that the data will be writen to.
        :param DataFrame df:
        :return:
        :rtype: CellsArea
        """
        start_row_idx = self.top_left_point.x
        start_col_idx = self.top_left_point.y
        if self.__with_header:
            if self.__header_direction == row:
                data_row_idx = start_row_idx + 1
                data_col_idx = start_col_idx
            else:
                data_row_idx = start_row_idx
                data_col_idx = start_col_idx + 1
        else:
            data_row_idx = start_row_idx
            data_col_idx = start_col_idx

        if self.__rows_limit is not None:
            num_of_data_rows_to_write = min(self.__rows_limit, df.shape[0])
        else:
            num_of_data_rows_to_write = df.shape[0]

        if row == self.__header_direction:
            bottom_right_point = Cell(
                data_row_idx + num_of_data_rows_to_write - 1,
                data_col_idx + len(self.__headers)
            )
        else:
            bottom_right_point = Cell(
                data_row_idx + len(self.__headers),
                data_col_idx + num_of_data_rows_to_write - 1
            )
        return CellsArea(
            self.top_left_point,
            bottom_right_point,
            sheet_name=self.__sheet_name
        )

    def parse_format_config(self, config):
        if "format" not in config:
            return None
        format_config = config["format"]
        format_dict = {}
        if "*" in format_config:
            format_str = format_config["*"]
            for header in self.__headers:
                format_dict[header] = format_str
        format_dict.update({
            k: v for k, v in format_config.items() if k != "*"
        })
        return {header: get_value_convert_func(format_str)
                for header, format_str in format_dict.items()}


class FileWriter(object):

    def __init__(self, template, config):
        """

        :param Template template:
        :param dict[str,dict[str,object]] config:
        """
        template_info_items = template.info_items

        self.__target_file_type = template.file_type

        info_item_dict = {}
        for info_item in template_info_items:
            if info_item.mapping_name in info_item_dict:
                err_msg = "Duplicated mapping name in the reader template file.\n Mapping name :" + info_item.mapping_name
                raise Exception(err_msg)
            info_item_dict[info_item.mapping_name] = info_item

        if self.__target_file_type == "excel":
            writer_config_dict = config["writers"]
            self.__writer_dict = {
                mapping_name: ExcelDataFrameWriter(info_item_dict[mapping_name], mapping_writer_config)
                for mapping_name, mapping_writer_config in writer_config_dict.items()
            }
        else:
            # in current version, csv file only support horizontal header.
            for info_item in template_info_items:
                if info_item.header_direction == column:
                    err_msg = "In current version, csv file only support horizontal header."
                    raise Exception(err_msg)
            writer_config_dict = config["writers"]
            self.__writer_dict = {
                mapping_name: ExcelDataFrameWriter(info_item_dict[mapping_name], mapping_writer_config)
                for mapping_name, mapping_writer_config in writer_config_dict.items()
            }

    def write(self, data_frame_dict, target_file_path):
        """

        :param dict[str,DataFrame] data_frame_dict:
        :return:
        """
        if self.__target_file_type == "excel":
            self.__check_write_area(data_frame_dict)

            for mapping_name, df in data_frame_dict.items():
                if mapping_name in self.__writer_dict:
                    df_writer = self.__writer_dict[mapping_name]
                    df_writer.write(df, target_file_path)
        else:
            df_to_write = self.__merge_data_into_single_df(data_frame_dict)
            df_to_write.to_csv(
                target_file_path, header=False, index=False,
            )

    def __merge_data_into_single_df(self, data_frame_dict):
        write_area_list = []
        for mapping_name, df in data_frame_dict.items():
            if mapping_name in self.__writer_dict:
                df_writer = self.__writer_dict[mapping_name]
                write_area_list.append(df_writer.get_write_area(df))
        max_x = sorted([area.bottom_right_point.x for area in write_area_list], reverse=True)[0]
        max_y = sorted([area.bottom_right_point.y for area in write_area_list], reverse=True)[0]
        empty_column = [None for x in range(0, max_x)]
        data_dict = {y: empty_column.copy() for y in range(0, max_y)}
        for mapping_name, df in data_frame_dict.items():
            if mapping_name in self.__writer_dict:
                df_writer = self.__writer_dict[mapping_name]
                df_to_write = df_writer.get_data_frame_to_write(df)
                based_x, based_y = df_writer.top_left_point.x, df_writer.top_left_point.y
                for dx in range(0, df_to_write.shape[0]):
                    for dy in range(0, df_to_write.shape[1]):
                        data_dict[based_y + dy][based_x + dx] = df_to_write.iloc[dx].iloc[dy]
        return DataFrame(
            data_dict
        )

    def __check_write_area(self, data_frame_dict):
        """

        :param dict[str,DataFrame] data_frame_dict:
        :return:
        """
        cells_areas = []
        for mapping_name, df in data_frame_dict.items():
            if mapping_name in self.__writer_dict:
                writer = self.__writer_dict[mapping_name]
                cells_areas.append(writer.get_write_area(df))

        max_idx = len(cells_areas)
        for i in range(0, max_idx):
            area_1 = cells_areas[i]
            for j in range(i + 1, max_idx):
                area_2 = cells_areas[j]
                if area_1.is_intersect_with(area_2):
                    err_msg = "The Area about to write will be intersected."
                    raise Exception(err_msg)
