import json
import os

import openpyxl
from pandas import DataFrame

from xl_transform.common import Template, TemplateInfoItem, row
from xl_transform.common.ConfigPaseUtils import get_value_convert_func
from xl_transform.common.model import CellsArea, Cell
from xl_transform.writer import ExcelDataWriter


class ExcelDataFrameWriter(object):

    def __init__(self, info_item, config=None):
        """

        :param TemplateInfoItem info_item:
        :param dict[str,object] config:
        """
        self.__sheet_name = info_item.sheet_name
        self.__top_left_point = info_item.top_left_point
        self.__mapping_name = info_item.mapping_name
        self.__headers = list(info_item.headers)
        self.__header_direction = info_item.header_direction
        self.__with_header = True if config is not None and "with_header" in config and "true" == config[
            "with_header"].lower() else False
        self.__rows_limit = int(config["rows_limit"]) if config is not None and "rows_limit" in config else None
        self.__format = self.parse_format_config(config) if config is not None else None

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

    def __init__(self, template, config=None):
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

        if config is not None and "writers" in config:
            writer_config_dict = config["writers"]
            self.__writer_dict = {
                mapping_name: ExcelDataFrameWriter(info_item, writer_config_dict[mapping_name])
                for mapping_name, info_item in info_item_dict.items()
            }
        else:
            self.__writer_dict = {
                mapping_name: ExcelDataFrameWriter(info_item)
                for mapping_name, info_item in info_item_dict.items()
            }

    def write_data(self, data_frame_dict, target_file_path):
        """

        :param dict[str,DataFrame] data_frame_dict:
        :param str target_file_path:
        :return:
        """
        self.__check_write_area(data_frame_dict)

        for mapping_name, df in data_frame_dict.items():
            if mapping_name in self.__writer_dict:
                df_writer = self.__writer_dict[mapping_name]
                df_writer.write(df, target_file_path)

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
                    err_msg = "The Area about to write_data will be intersected."
                    raise Exception(err_msg)

    @staticmethod
    def write(
            target_path,
            template_path,
            data,
            config_path=None
    ):
        """

        :param str target_path: source_path: The path of the target excel file.
        :param str template_path: The path of the template excel file.
        :param dict[str,DataFrame] data: The data to write_data.
        :param str config_path: The path of the configuration file.
        :return:
        """
        parent_dir_of_target = os.path.dirname(target_path)
        # check target path
        if not os.path.exists(parent_dir_of_target):
            err_msg = "The parent folder of the target path does not exists, please create '{}' first.".format(
                parent_dir_of_target
            )
            raise Exception(err_msg)
        elif not os.path.isdir(parent_dir_of_target):
            err_msg = "The parent path '{}' is not a folder, please choose other folder.".format(
                parent_dir_of_target
            )
            raise Exception(err_msg)

        # check template path
        if not os.path.exists(template_path):
            err_msg = "The given path of the template file does not exists, please check you path: '{}'.".format(
                template_path
            )
            raise Exception(err_msg)
        elif not os.path.isfile(template_path):
            err_msg = "The given path of the template file is not a file, please check you path: '{}'.".format(
                template_path
            )
            raise Exception(err_msg)
        template = Template(template_path)

        # check config path
        if config_path is not None:
            if not os.path.exists(config_path):
                err_msg = "The given path of the config file is not a file, please check you path: '{}'.".format(
                    template_path
                )
                raise Exception(err_msg)
            elif not os.path.isfile(config_path):
                err_msg = "The given path of the config file is not a file, please check you path: '{}'.".format(
                    template_path
                )
                raise Exception(err_msg)
            with open(config_path, 'r') as config_file:
                config = json.load(config_file)
        else:
            config = None

        FileWriter(template, config).write_data(data, target_path)
