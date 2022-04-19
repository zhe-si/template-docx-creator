import os


def make_sure_path(path: str):
    """
    Make sure a path exists.
    """
    if not os.path.exists(path):
        os.makedirs(path)
