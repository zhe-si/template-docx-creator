import json
from abc import ABCMeta, abstractmethod

from exception import JsonError
from helper.os_helper import *


class DataLoader(metaclass=ABCMeta):
    @abstractmethod
    def load_data(self) -> [dict, None]:
        """加载单组数据，返回内容标签名称和内容，多次加载分别返回每组数据，若无数据，返回 None"""
        pass


class StaticDataLoader(DataLoader):
    """静态数据加载器，存储所有数据并依次加载"""

    def __init__(self, datas: [list[dict], dict]):
        if isinstance(datas, dict):
            datas = [datas]
        self._datas = datas
        self._datas_len = len(self._datas)
        self._index = 0

    def load_data(self) -> [dict, None]:
        if self._index >= self._datas_len:
            return None
        data = self._datas[self._index]
        self._index += 1
        return data


class JsonDataLoader(DataLoader):
    """
    json数据加载器，解析后存储在内存中依次加载
    支持从文件加载、字符串加载，若有多组数据源，优先加载字符串，其他数据源懒加载
    """

    def __init__(self, file_path=None, json_str=None):
        self._file_path = file_path
        self._need_load_file = True if self._file_path else False
        self._json_str = json_str

        self._data_loader = None

        if self._json_str is not None:
            try:
                json_data = json.loads(self._json_str)
            except json.JSONDecodeError as e:
                raise JsonError(f'解析json字符串失败，\n\t字符串：{self._json_str}\n\t错误：{str(e)}')
            self._add_datas(json_data)
        elif self._need_load_file:
            self._load_file_data()
            self._need_load_file = False
        else:
            self._data_loader = StaticDataLoader([])

    def _load_file_data(self):
        if self._file_path is not None:
            with open(self._file_path, 'r', encoding=check_file_coding(self._file_path)) as f:
                try:
                    json_data = json.load(f)
                except json.JSONDecodeError as e:
                    raise JsonError(f'解析json文件 {self._file_path} 失败，错误：{str(e)}')
            self._add_datas(json_data)

    def _add_datas(self, json_data):
        json_datas = []
        if isinstance(json_data, dict):
            json_datas.append(json_data)
        elif isinstance(json_data, list):
            json_datas += json_data
        else:
            raise JsonError('json数据格式错误，应为单组插入内容的字典或多组插入内容的列表')
        self._data_loader = StaticDataLoader(json_datas)

    def load_data(self) -> [dict, None]:
        data = self._data_loader.load_data()
        if data is None:
            if self._need_load_file:
                self._load_file_data()
                self._need_load_file = False
                return self._data_loader.load_data()
            else:
                return None
        return data
