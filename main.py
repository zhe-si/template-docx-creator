import re
from enum import Enum, unique

from docx import Document

import labels
from helper.os_helper import *


_content_label_re = re.compile(r'{{(.*?)}}')

registered_labels = {label.get_type(): label for label in labels.LabelManager.get_labels()}

insert_point_content_types = [label.get_type() for label in registered_labels.values() if label.has_content()]
insert_point_no_content_types = [label.get_type() for label in registered_labels.values() if not label.has_content()]
insert_point_types = insert_point_no_content_types + insert_point_content_types


def match(file_path: str, save_path: str, datas: dict):
    # 更新全局的标签信息
    # update_labels_info()

    # 注册每次模板生成过程的静态插入数据
    static_datas = {}
    for label in registered_labels.values():
        label.register_static_datas(static_datas)

    # 校验模板并在预处理阶段插入无内容标签的信息
    def insert_data_to_no_content_point(p_d: dict):
        p_t = p_d['type']
        if p_t in insert_point_no_content_types:
            no_content_label = registered_labels[p_t]
            no_content_label.insert_data_to_point(p_d, None, static_datas)
            return True
        return False

    # 模板检查与预处理
    check_result = check_template(file_path, insert_data_to_no_content_point)
    # 模板校验失败处理
    if check_result['code'].is_error():
        print(f'模板校验失败\n\t错误代码：{check_result["code"]}\n\t错误信息：{check_result["msg"]}')
        return
    insert_points = check_result['data']['insert_points']
    document = check_result['data']['document']

    no_data_points = {}

    # 处理有内容类型插入点，检查并插入数据
    solve_content_labels(insert_points, datas, static_datas, no_data_points)

    document.save(save_path)

    # 打印无数据对应的内容标签信息
    print_no_data_points(no_data_points)


def solve_content_labels(insert_points, datas, static_datas, no_data_points):
    for point_name, point_data in insert_points.items():
        if point_name in datas:
            data = datas[point_name]
            label = registered_labels[point_data['type']]
            if not label.check_data_type(data):
                print(f"插入内容类型 {point_data['type']} 不能匹配数据 {type(data)}。内容标签为 {point_data['text']}，原文为 {point_data['run'].text}")
                continue
            label.insert_data_to_point(point_data, data, static_datas)
        else:
            no_data_points[point_name] = point_data


def update_labels_info():
    global registered_labels, insert_point_no_content_types, insert_point_content_types, insert_point_types
    registered_labels = {label.get_type(): label for label in labels.LabelManager.get_labels()}
    insert_point_content_types = [label.get_type() for label in registered_labels.values() if label.has_content()]
    insert_point_no_content_types = [label.get_type() for label in registered_labels.values() if not label.has_content()]
    insert_point_types = insert_point_no_content_types + insert_point_content_types


@unique
class CheckCode(Enum):
    SUCCESS = 1
    # 严重错误
    NAME_REPEAT = -1
    LABEL_FORMAT_ERROR = -2

    def __str__(self):
        return f'{super(CheckCode, self).__str__()} {self.value}'

    def is_error(self):
        return self.value < 0


def check_template(file_path: str, insert_operation: callable) -> dict:
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
            run_insert_points = _content_label_re.findall(p_run.text)
            for point in run_insert_points:
                # 内容标签格式检查
                point_split = point.split(':')
                if len(point_split) != 2:
                    return {"code": CheckCode.LABEL_FORMAT_ERROR,
                            "msg": f"内容标签'{point}'格式错误，应该为'{{label_type:label_name}}'，label_name 可省略，但不可省略 ':'",
                            "data": None}
                point_type, point_name = point_split
                # 忽略无法识别类型的内容标签
                if point_type not in insert_point_types:
                    continue

                point_data = {'name': point_name, 'type': point_type, 'text': '{{' + point + '}}',
                              'run': p_run, 'run_index': p_run_index, 'paragraph': paragraph, 'document': document}

                # 插入点信息处理
                if not insert_operation(point_data):
                    if point_name in insert_points:
                        return {"code": CheckCode.NAME_REPEAT,
                                "msg": f"插入点名称 '{point_name}' 在模板内重复，内容标签为'{point_data['text']}'，原文为'{point_data['run'].text}'",
                                "data": None}
                    insert_points[point_name] = point_data

    return {"code": CheckCode.SUCCESS,
            "msg": "successful",
            "data": {"document": document, "insert_points": insert_points}}


def print_no_data_points(no_data_points):
    if len(no_data_points) == 0:
        return
    print("无数据对应的内容标签：")
    i = 1
    for point_name, point_data in no_data_points.items():
        print(f"  ({i}) 标签名为'{point_name}'、类型为'{point_data['type']}'的内容标签'{point_data['text']}'无法匹配到数据，原文：{point_data['run'].text}")


def main():
    datas = {
        'file_title0': '插入文档的标题',
        'file_content0': '插入文档的内容，' * 20,
        'file_content1': '文字的样式保持一致',
        'file_ul1': ['无序列表内容1', '无序列表内容2', '无序列表内容3', ],
        'file_ol1': ['有序列表内容1', '有序列表内容1', '有序列表内容1', ],
        'file_link1': ('百度的链接', 'https://www.baidu.com'),
        'file_img1': ('这是图片的介绍', 'data/p1.jpg'),
        'file_table1': [
            ['姓名', '学号', '性别', '班级'],
            ['张三', '123456789', '男', '计算机科学与技术1401'],
            ['李四', '987654321', '男', '软件工程'],
            ['小红', '77777', '女', '软件工程'],
        ],
        '123': '模板中不存在名字为123的内容标签，该数据自动被忽略',
    }

    tem_path = "data/template-demo.docx"
    to_dir = "test_data"
    word_name = "word.docx"
    make_sure_path(to_dir)

    match(tem_path, f"{to_dir}/{word_name}", datas)


if __name__ == '__main__':
    main()
