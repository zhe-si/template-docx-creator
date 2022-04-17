import re

from docx import Document

import labels


def match(file_path: str, save_path: str, datas: dict):
    registered_labels = {label.get_type(): label for label in labels.LabelManager.get_labels()}

    insert_point_content_types = [label.get_type() for label in registered_labels.values() if label.has_content()]
    insert_point_no_content_types = [label.get_type() for label in registered_labels.values() if not label.has_content()]
    insert_point_types = insert_point_no_content_types + insert_point_content_types

    static_datas = {}

    content_label_re = re.compile(r'{{(.*?)}}')

    document = Document(file_path)

    insert_points = {}
    for paragraph in document.paragraphs:
        for p_run in paragraph.runs:
            insert_points = content_label_re.findall(p_run.text)
            for point in insert_points:
                point_type, point_name = point.split(':')
                if point_type not in insert_point_types:
                    raise Exception("Invalid insert point type: " + point_type)
                # 处理无内容类型
                if point_type in insert_point_no_content_types:
                    point_data = {'name': point_name, 'type': point_type, 'run': p_run, 'text': '{{' + point + '}}'}
                    registered_labels[point_type].insert_data_to_point(point_data, None, static_datas)
                # 记录有内容类型插入点
                else:
                    if point_name in insert_points:
                        raise Exception("Duplicate insert point: " + point_name)
                    insert_points[point_name] = {'name': point_name, 'type': point_type, 'run': p_run, 'text': '{{' + point + '}}'}

    no_data_points = {}

    for point_name, point_data in insert_points.items():
        if point_name in datas:
            registered_labels[point_data['type']].insert_data_to_point(point_data, datas[point_name], static_datas)
        else:
            no_data_points[point_name] = point_data

    document.save(save_path)

    print(no_data_points)


def main():
    datas = {
        't1': 'tt1',
        't2': 'tt2',
        't3': 'tt3',
        't4': 'tt4',
    }

    match("data/t2.docx", "data/r2.docx", datas)


if __name__ == '__main__':
    main()
