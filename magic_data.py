import sys


class MagicData:
    """魔法数据"""
    _magic_data_list = ['__doc_name__', '__tem_name__']

    def __init__(self, datas: dict):
        self._datas = datas

    def get_doc_name(self, default=None) -> [str, None]:
        """获取生成文档的名称"""
        doc_name = self._datas.get('__doc_name__', None)
        if doc_name is None:
            doc_name = default
        else:
            if not doc_name.endswith('.docx'):
                doc_name = doc_name + '.docx'
        return doc_name

    @classmethod
    def _set_data(cls, data, key: str, value, is_cover):
        """向数据设置魔术数据"""
        key = '_'.join(key.split('_')[-2:])
        key = f'__{key}__'
        if is_cover or key not in data:
            data[key] = value

    @classmethod
    def set_doc_name(cls, data, tem_name, is_cover=False):
        """向数据设置要使用的模板的名称"""
        cls._set_data(data, sys._getframe().f_code.co_name, tem_name, is_cover)
        return True

    def get_tem_name(self, default=None) -> [str, None]:
        """获取要使用的模板的名称"""
        return self._datas.get('__tem_name__', default)

    @classmethod
    def set_tem_name(cls, data, tem_name, is_cover=False):
        """向数据设置要使用的模板的名称"""
        cls._set_data(data, sys._getframe().f_code.co_name, tem_name, is_cover)
        return True

    @classmethod
    def check_all_magic_data(cls, data):
        """检查数据是否包含所有魔术数据"""
        for magic_data in cls._magic_data_list:
            if magic_data not in data:
                return False
        return True
