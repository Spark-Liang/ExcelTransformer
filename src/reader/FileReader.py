from pandas import DataFrame

from common import Template


class FileReader(object):

    def __init__(self, template, config):
        """

        :param Template template:
        :param dict[str,object] config:
        """
        self.__template = template
        self.__config = config

    def read(self, filepath):
        """

        :param filepath:
        :return:
        :rtype: dict[str,DataFrame]
        """
        pass
