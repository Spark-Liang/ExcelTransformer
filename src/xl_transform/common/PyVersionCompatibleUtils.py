import sys


def is_string(val):
    return (sys.version_info.major == 2 and isinstance(val, unicode)) \
           or (sys.version_info.major == 3 and isinstance(val, str))


from datetime import datetime


class DateTimeConvertUtils(object):

    @staticmethod
    def timestamp_to_datetime(ts):
        """

        :param ts:
        :return:
        :rtype: datetime
        """
        datetime.fromtimestamp(ts)
