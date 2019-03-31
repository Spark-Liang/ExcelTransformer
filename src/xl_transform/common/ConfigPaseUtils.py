import decimal
import re
from datetime import datetime

from openpyxl.cell.cell import Cell


def get_cell_value_convert_func(type_hint):
    """

    :param str type_hint:
    :return:
    """

    type_date_pattern = r"date\((.+)\)"
    type_decimal_pattern = r"decimal\((\d+)\)"

    if re.match(type_date_pattern, type_hint):
        date_format = re.match(type_date_pattern, type_hint).group(1)

        def tran_func(cell):
            """

            :param Cell cell:
            :return:
            """
            if isinstance(cell.value, str):
                return datetime.strptime(cell.value, date_format)
            return cell.value

        return tran_func

    elif re.match(type_decimal_pattern, type_hint):
        prec = re.match(type_decimal_pattern, type_hint).group(1)
        formatter = "{:." + prec + "f}"

        def trans_func(cell):
            """

            :param Cell cell:
            :return:
            """
            val = cell.value
            if isinstance(val, str):
                return decimal.Decimal(cell.value)
            elif isinstance(val, (int, float)):
                return decimal.Decimal(formatter.format(val))
            return cell.value

        return trans_func
    return None


def get_value_convert_func(format_str):
    """

    :param str format_str:
    :return:
    """

    type_date_pattern = r"date\((.+)\)"
    type_decimal_pattern = r"decimal\((\d+)\)"

    if re.match(type_date_pattern, format_str):
        date_format = re.match(type_date_pattern, format_str).group(1)

        def trans_func(val):
            if isinstance(val, datetime):
                return datetime.strftime(val, date_format)
            return val

        return trans_func
    elif re.match(type_decimal_pattern, format_str):
        prec = re.match(type_decimal_pattern, format_str).group(1)
        format_str = r"{:." + prec + r"f}"

        def trans_func(val):
            if isinstance(val, (int, float, decimal.Decimal)):
                return format_str.format(val)
            return val

        return trans_func
    return None
