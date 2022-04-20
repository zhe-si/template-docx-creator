import re
from enum import Enum, unique

from docx import Document

import labels


def is_no_content_point(p_d):
    p_t = p_d['type']
    if p_t in TemplateAnalyzer.insert_point_no_content_types:
        return True
    return False


class TemplateAnalyzer:
    _content_label_re = re.compile(r'{{(.*?)}}')

    registered_labels = {label.get_type(): label for label in labels.LabelManager.get_labels()}

    insert_point_content_types = [label.get_type() for label in registered_labels.values() if label.has_content()]
    insert_point_no_content_types = [label.get_type() for label in registered_labels.values() if
                                     not label.has_content()]
    insert_point_types = insert_point_no_content_types + insert_point_content_types

    static_datas = {}

    @classmethod
    def update_labels_info(cls):
        cls.registered_labels = {label.get_type(): label for label in labels.LabelManager.get_labels()}
        cls.insert_point_content_types = [label.get_type() for label in cls.registered_labels.values() if label.has_content()]
        cls.insert_point_no_content_types = [label.get_type() for label in cls.registered_labels.values() if not label.has_content()]
        cls.insert_point_types = cls.insert_point_no_content_types + cls.insert_point_content_types

    @classmethod
    def register_static_datas(cls):
        """标签依次向 static_datas 注册静态数据，每次注册会清空之前的数据"""
        cls.static_datas.clear()
        for label in cls.registered_labels.values():
            label.register_static_datas(cls.static_datas)

    @unique
    class CheckCode(Enum):
        SUCCESS = 1
        # 严重错误
        NAME_REPEAT = -1  # 标签名重复
        LABEL_FORMAT_ERROR = -2  # 标签格式错误
        NAME_EMPTY = -3  # 标签名为空

        def __str__(self):
            return f'{super(TemplateAnalyzer.CheckCode, self).__str__()} {self.value}'

        def is_error(self):
            return self.value < 0

    @staticmethod
    def _create_check_info(code: CheckCode, msg: str = '', data: dict = None):
        return {'code': code, 'msg': msg, 'data': data}

    @classmethod
    def check_template(cls, file_path: str, insert_operation: callable = is_no_content_point) -> dict:
        """
        校验出错返回错误信息，成功返回插入点信息字典与文件

        :param file_path: 文件路径
        :param insert_operation: 插入操作，是一个可调用对象（包括函数），参数为插入点信息字典，返回bool表示是否拦截该内容标签，若不拦截，保存到插入点信息字典

        :return: (code, msg, data) 三元组。若 code = 1 {CheckCode.SUCCESS} 表示成功，data 为字典，包含 insert_points:插入点信息字典 与 document:文件对象；若 code < 0 表示失败，msg 为错误信息，data 为 None。
                 具体错误码说明：
                 -1 {CheckCode.NAME_REPEAT} [严重]: 内容标签名称重复。
                 -2 {CheckCode.LABEL_FORMAT_ERROR} [严重]: 内容标签格式错误。
        """
        document = Document(file_path)
        insert_points = {}

        for paragraph in document.paragraphs:
            for p_run_index, p_run in enumerate(paragraph.runs):
                run_insert_points = cls._content_label_re.findall(p_run.text)
                for point in run_insert_points:
                    # 内容标签格式检查
                    point_split = point.split(':')
                    if len(point_split) != 2:
                        return cls._create_check_info(cls.CheckCode.LABEL_FORMAT_ERROR, f"内容标签'{point}'格式错误，应该为'{{label_type:label_name}}'，label_name 可省略，但不可省略 ':'")
                    point_type, point_name = point_split
                    # 忽略无法识别类型的内容标签
                    if point_type not in cls.insert_point_types:
                        continue

                    point_data = {'name': point_name, 'type': point_type, 'text': '{{' + point + '}}',
                                  'run': p_run, 'run_index': p_run_index, 'paragraph': paragraph, 'document': document}

                    # 插入点信息处理
                    if not insert_operation(point_data):
                        if len(point_name) == 0:
                            return cls._create_check_info(cls.CheckCode.NAME_EMPTY, f"插入点名称为空，内容标签为'{point_data['text']}'，原文为'{point_data['run'].text}'")
                        if point_name in insert_points:
                            return cls._create_check_info(cls.CheckCode.NAME_REPEAT, f"插入点名称 '{point_name}' 在模板内重复，内容标签为'{point_data['text']}'，原文为'{point_data['run'].text}'")
                        insert_points[point_name] = point_data

        return cls._create_check_info(cls.CheckCode.SUCCESS, 'successful', {'insert_points': insert_points, 'document': document})

    @staticmethod
    def print_check_info(check_info: dict, show_detail=False):
        """
        打印校验信息

        :param check_info: 校验信息
        :param show_detail: 打印详细信息
        """
        if check_info['code'] == TemplateAnalyzer.CheckCode.SUCCESS:
            insert_points = check_info['data']['insert_points']
            print(f"模板校验成功，共有 {len(insert_points)} 个插入点")
            if show_detail:
                for i, p_n in enumerate(insert_points):
                    p_d = insert_points[p_n]
                    print(f'\t{i + 1}、{p_n}：{p_d["text"]}')
        else:
            print(f'模板校验失败\n\t错误代码：{check_info["code"]}\n\t错误信息：{check_info["msg"]}')
