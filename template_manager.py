import json
import os.path
import shutil

from helper.os_helper import *
from template_analyzer import TemplateAnalyzer


class TemplateManager:
    _templates = {}
    _DEFAULT_SAVE_PATH = './templates/'
    _DEFAULT_INFO_PATH = './templates/.index'

    @classmethod
    def add_template(cls, template_path):
        """
        添加模板
        该路径模板已存在，则返回名字；不存在则校验模板，成功则添加并返回名字，失败返回None
        """
        if not template_path.endswith('.docx') or not os.path.isfile(template_path):
            return None
        file_name = get_file_name(template_path, has_suffix=False)

        tem_n = cls.get_template_name_by_path(template_path)
        if tem_n is None:
            # 不允许模板重名
            if file_name in cls._templates:
                return None
            if cls.check_template(template_path):
                cls._templates[file_name] = os.path.abspath(template_path)
                tem_n = file_name
        return tem_n

    @classmethod
    def add_template_with_copy(cls, template_path, save_dir=_DEFAULT_SAVE_PATH):
        if not template_path.endswith('.docx') or not os.path.isfile(template_path):
            return None
        file_name = os.path.basename(template_path)
        make_sure_path(save_dir)
        tem_path = '/'.join([save_dir, file_name])
        if os.path.abspath(tem_path) != os.path.abspath(template_path):
            shutil.copyfile(template_path, tem_path)
        return cls.add_template(tem_path)

    @classmethod
    def get_template_path(cls, name):
        if name in cls._templates:
            return cls._templates[name]
        return None

    @classmethod
    def get_template_name_by_path(cls, tem_path):
        if not os.path.exists(tem_path):
            return None
        for name, path in cls._templates.items():
            if os.path.samefile(path, tem_path):
                return name
        return None

    @classmethod
    def get_template_list(cls):
        return list(cls._templates.keys())

    @classmethod
    def rename_template(cls, name, new_name):
        if name in cls._templates and new_name not in cls._templates:
            tem_path = cls._templates[name]
            new_path = '/'.join([os.path.dirname(tem_path), f'{new_name}.docx'])
            if os.path.exists(new_path):
                return False
            shutil.copyfile(tem_path, new_path)
            cls._templates[new_name] = new_path
            cls._templates.pop(name)
            return True
        return False

    @classmethod
    def remove_template(cls, name):
        if name in cls._templates:
            cls._templates.pop(name)
            return True
        return False

    @classmethod
    def load_info(cls, info_path=_DEFAULT_INFO_PATH):
        if os.path.exists(info_path):
            with open(info_path, 'r') as f:
                info = json.load(f)
            for name, path in info.items():
                cls.add_template(path)

    @classmethod
    def save_info(cls, info_path=_DEFAULT_INFO_PATH):
        make_sure_path(os.path.dirname(info_path))
        with open(info_path, 'w') as f:
            json.dump(cls._templates, f)

    @classmethod
    def load_path(cls, path=_DEFAULT_SAVE_PATH):
        if os.path.isdir(path):
            for f_n in os.listdir(path):
                tem_path = '/'.join([path, f_n])
                if os.path.isfile(tem_path):
                    cls.add_template(tem_path)

    @classmethod
    def load_default(cls):
        cls.load_info()
        cls.load_path()

    @classmethod
    def print_template_info(cls):
        print("当前模板列表：")
        i = 1
        for name, path in cls._templates.items():
            print(f'\t{i}. {name}: {path}')
            i += 1

    @staticmethod
    def check_template(tem_path: str, print_info=False):
        """
        根据模板检查模板是否符合要求

        :param tem_path: 模板文件路径
        :param print_info: 是否打印检查过程的信息，默认开启详细打印模式
        :return: 模板是否符合要求
        """
        check_result = TemplateAnalyzer.check_template(tem_path)
        if print_info:
            TemplateAnalyzer.print_check_info(check_result, show_detail=True)
        # 模板校验失败直接退出
        if check_result['code'].is_error():
            return False
        return True
