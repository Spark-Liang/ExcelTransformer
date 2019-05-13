def format_collection(collection):
    """

    :param dict or set or list or tuple collection:
    :return:
    """
    if is_collection(collection):
        if isinstance(collection, set):
            return format_set(collection)
        if isinstance(collection, list):
            return format_list(collection)
        if isinstance(collection, dict):
            return format_dict(collection)
    else:
        return str(collection)



def is_collection(o):
    return (isinstance(o, set) or isinstance(o, tuple)
            or isinstance(o, list) or isinstance(o, dict))


def format_set(set_to_format):
    return r"{" + ",".join([
        format_collection(x) if is_collection(x) else str(x) for x in set_to_format
    ]) + r"}"


def format_tuple(tuple_to_format):
    return r"(" + ",".join([
        format_collection(x) if is_collection(x) else str(x) for x in tuple_to_format
    ]) + r")"


def format_list(list_to_format):
    return r"[" + ",".join([
        format_collection(x) if is_collection(x) else str(x) for x in list_to_format
    ]) + r"]"


def format_dict(dict_to_format):
    """

    :param dict dict_to_format:
    :return:
    """
    return r"{" + ",\n".join([
        "{} : {}".format(str(x), format_collection(y)) if is_collection(y) else "{} : {}".format(str(x), str(y))
        for x, y in dict_to_format.items()
    ]) + r"}"
