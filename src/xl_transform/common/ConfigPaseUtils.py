import datetime
import decimal
import numbers
import re
import time


def get_str2type_transfer_func(type_hint):
    """

    :param str type_hint:
    :return:
    """

    type_date_pattern = r"date\((.+)\)"
    type_decimal_pattern = r"decimal\((\d+)\)"

    if re.match(type_date_pattern, type_hint):
        date_format = re.match(type_date_pattern, type_hint).group(1)

        def tran_func(val):
            if isinstance(val, str):
                struct_time = time.strptime(val, date_format)
                return datetime.datetime(struct_time)
            return val

        return tran_func

    elif re.match(type_decimal_pattern, type_hint):
        prec = re.match(type_decimal_pattern, type_hint).group(1)
        context = decimal.getcontext()
        context.prec = int(prec)

        def trans_func(val):
            if isinstance(val, str):
                return decimal.Decimal(val, context)
            return val

        return trans_func
    return None


def get_type2str_transfer_func(type_hint):
    """

    :param str type_hint:
    :return:
    """

    type_date_pattern = r"date\((.+)\)"
    type_decimal_pattern = r"decimal\((\d+)\)"

    if re.match(type_date_pattern, type_hint):
        date_format = re.match(type_date_pattern, type_hint).group(1)

        def trans_func(val):
            return time.strftime(date_format, val)

        return trans_func
    elif re.match(type_decimal_pattern, type_hint):
        prec = re.match(type_decimal_pattern, type_hint).group(1)
        format_str = r"{:." + prec + r"f}"

        def trans_func(val):
            if isinstance(val, numbers.Number):
                return format_str.format(val)
            return val

        return trans_func
    return None
