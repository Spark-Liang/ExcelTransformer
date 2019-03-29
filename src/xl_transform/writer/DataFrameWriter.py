from pandas import DataFrame


class AbstractDataFrameWriter(object):

    def write(self, df):
        """
        Write the given DataFrame into target area.
        :param DataFrame df :
        :return:
        """
        pass

    def get_write_area(self, df):
        """
        Get the area that the data will be writen to.
        :param DataFrame df:
        :return:
        :rtype: CellsArea
        """
        pass
