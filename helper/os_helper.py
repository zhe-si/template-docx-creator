import os

import chardet


def make_sure_path(path: str):
    """
    Make sure a path exists.
    """
    if not os.path.exists(path):
        os.makedirs(path)


def check_file_coding(file_path: str):
    with open(file_path, 'rb') as f:
        return chardet.detect(f.read(100))['encoding']


def get_file_name(file_path: str, has_suffix=True):
    p = os.path.basename(file_path)
    if not has_suffix:
        sp = p.split('.')
        sp = sp[:-1] if len(sp) > 1 else sp
        p = '.'.join(sp)
    return p


def get_file_suffix(file_path):
    return file_path.split('.')[-1]


def get_file_name_no_repeat(file_path):
    file_name = get_file_name(file_path, has_suffix=False)
    file_suffix = get_file_suffix(file_path)
    if not os.path.exists(file_path):
        return f'{file_name}.{file_suffix}'
    base_dir = os.path.dirname(file_path)
    i = 1
    while True:
        n = f'{file_name}_{i}.{file_suffix}'
        if not os.path.exists(f'{base_dir}/{n}'):
            return n
