from typing import Iterable


def check_iterable_type(obj, ele_type=None) -> bool:
    return isinstance(obj, Iterable) and (not ele_type or all(isinstance(item, str) for item in obj))
