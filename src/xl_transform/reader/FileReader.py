from pandas import DataFrame

from xl_transform.reader import DataFrameReader


class FileReader(object):

    def __init__(self, template, config):
        """

        :param Template template:
        :param dict[str,dict[str,object]] config:
        """
        template_info_items = template.info_items
        info_item_dict = {}
        for info_item in template_info_items:
            if info_item.mapping_name in info_item_dict:
                err_msg = "Duplicated mapping name in the reader template file.\n Mapping name :" + info_item.mapping_name
                raise Exception(err_msg)
            info_item_dict[info_item.mapping_name] = info_item

        reader_config_dict = config["readers"]
        self.__reader_dict = {
            mapping_name: DataFrameReader(info_item_dict[mapping_name], mapping_reader_config)
            for mapping_name, mapping_reader_config in reader_config_dict.items()
        }

    def read(self, filepath):
        """

        :param filepath:
        :return:
        :rtype: dict[str,DataFrame]
        """
        result = {}
        for mapping_name, reader in self.__reader_dict.items():
            result[mapping_name] = reader.read(filepath)[1]
        return result
