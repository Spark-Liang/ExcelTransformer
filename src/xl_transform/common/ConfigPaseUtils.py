import decimal
import re
from datetime import datetime

from openpyxl.cell.cell import Cell
from openpyxl.styles import numbers


def get_cell_value_convert_func(type_hint):
    """

    :param str type_hint:
    :return:
    :rtype: None or (Cell) -> Any
    """

    type_str_pattern = r"str\((.+)\)"
    type_date_pattern = r"date\((.+)\)"
    type_decimal_pattern = r"decimal\((\d+)\)"

    if re.match(type_str_pattern, type_hint):
        str_format = re.match(type_str_pattern, type_hint).group(1)

        def trans_func(val):
            return str_format.format(val)

        return trans_func
    elif re.match(type_date_pattern, type_hint):
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


def get_format_func(format_str):
    """

    :param str format_str:
    :return:
    :rtype: None or (Cell,value) -> None
    """

    type_str_pattern = r"str\((.+)\)"
    type_date_pattern = r"date\((.+)\)"
    type_decimal_pattern = r"decimal\((\d+)\)"

    if re.match(type_str_pattern, format_str):
        str_format = re.match(type_str_pattern, format_str).group(1)

        def format_func(cell, val):
            cell.number_format = numbers.FORMAT_TEXT
            try:
                cell.value = str_format.format(val)
            except ValueError as e:
                err_msg = "Failed to use '{}' to format value {}\nReason is :{}".format(
                    str_format, val,
                    e
                )
                raise Exception(err_msg)

        return format_func
    elif re.match(type_date_pattern, format_str):
        date_format = re.match(type_date_pattern, format_str).group(1)
        xl_format_str = __convert_python_date_format_to_xl_date_format(date_format)

        def format_func(cell, val):
            cell.number_format = xl_format_str
            cell.value = val

        return format_func
    elif re.match(type_decimal_pattern, format_str):
        prec = re.match(type_decimal_pattern, format_str).group(1)
        prec_num = int(prec)
        xl_format_str = "0.{}_".format("0" * prec_num)

        def format_func(cell, val):
            cell.number_format = xl_format_str
            cell.value = val

        return format_func
    return None


def __convert_python_date_format_to_xl_date_format(format_str):
    # process year
    result = format_str.replace(r'%Y', 'yyyy')
    result = result.replace(r'%y', 'yy')

    # process month
    result = result.replace(r'%m', 'mm')
    result = result.replace(r'%b', 'mmm')

    # process day
    result = result.replace(r'%d', 'dd')

    # process hour
    result = result.replace(r'%H', 'hh')
    result = result.replace(r'%I', 'hh')

    # process minute
    result = result.replace(r'%M', 'mm')

    # process second
    result = result.replace(r'%S', 'ss')

    # process space
    result = result.replace(r' ', r'\ ')

    return result
