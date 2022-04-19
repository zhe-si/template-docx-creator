from abc import ABCMeta, abstractmethod


class DataLoader(metaclass=ABCMeta):
    @abstractmethod
    def load_data(self) -> [dict, None]:
        """加载单组数据，返回内容标签名称和内容，多次加载分别返回每组数据，若无数据，返回 None"""
        pass


class StaticDataLoader(DataLoader):
    """静态数据加载器，存储所有数据并依次加载"""
    def __init__(self, datas: list[dict]):
        self._datas = datas
        self._index = 0

    def load_data(self) -> [dict, None]:
        if self._index >= len(self._datas):
            return None
        data = self._datas[self._index]
        self._index += 1
        return data
