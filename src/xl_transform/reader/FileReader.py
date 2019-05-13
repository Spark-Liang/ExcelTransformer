import json
import os

from pandas import DataFrame

from xl_transform.common import Template
from xl_transform.reader import DataFrameReader


class FileReader(object):

    def __init__(self, template, config=None):
        """

        :param Template template:
        :param dict[str,dict[str,object]] config:
        """
        template_info_items = template.mapping_info_items
        info_item_dict = {}
        for info_item in template_info_items:
            if info_item.mapping_name in info_item_dict:
                err_msg = "Duplicated mapping name in the reader template file.\n Mapping name :{}".format(
                    info_item.mapping_name
                )
                raise Exception(err_msg)
            info_item_dict[info_item.mapping_name] = info_item

        if config is not None and "readers" in config:
            reader_config_dict = config["readers"]
            self.__reader_dict = {
                mapping_name: DataFrameReader(info_item, reader_config_dict[mapping_name])
                for mapping_name, info_item in info_item_dict.items()
            }
        else:
            self.__reader_dict = {
                mapping_name: DataFrameReader(info_item)
                for mapping_name, info_item in info_item_dict.items()
            }

    def read_data(self, filepath):
        """

        :param filepath:
        :return:
        :rtype: dict[str,DataFrame]
        """
        result = {}
        for mapping_name, reader in self.__reader_dict.items():
            result[mapping_name] = reader.read(filepath)[1]
        return result

    @staticmethod
    def read(
            source_path,
            template_path,
            config_path=None
    ):
        """

        :param str source_path: The path of the source excel file.
        :param str template_path: The path of the template excel file.
        :param str config_path: The path of the configuration file.
        :return:
        :rtype: dict[str,DataFrame]
        """
        # check target path
        if not os.path.exists(source_path):
            err_msg = "The given path of the source file does not exists, please check you input :'{}'.".format(
                source_path
            )
            raise Exception(err_msg)
        elif not os.path.isfile(source_path):
            err_msg = "The given path of the source file is not a file, please check you input :'{}'.".format(
                source_path
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

        return FileReader(template, config).read_data(source_path)
