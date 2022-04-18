import re
import time

from docx import Document


def main():
    insert_point_content_types = ['text', 'unordered-list', 'ordered-list', 'image', 'table', 'link']
    insert_point_no_content_types = ['date', 'time']
    insert_point_types = insert_point_no_content_types + insert_point_content_types

    static_datas = {}

    now_time = time.time()
    now_time_struct = time.localtime(now_time)
    date_s = time.strftime("%Y-%m-%d", now_time_struct)
    time_s = time.strftime("%H:%M:%S", now_time_struct)
    static_datas['date'] = date_s
    static_datas['time'] = time_s

    m = re.compile(r'{{(.*?)}}')

    document = Document("data/t2.docx")
    insert_points = {}
    for paragraph in document.paragraphs:
        for r in paragraph.runs:
            ps = m.findall(r.text)
            for p in ps:
                p_t, p_n = p.split(':')
                if p_t not in insert_point_types:
                    raise Exception("Invalid insert point type: " + p_t)
                # 处理无内容类型
                if p_t in insert_point_no_content_types:
                    t = r.text.replace('{{' + p + '}}', static_datas[p_t])
                    r.text = t
                # 记录有内容类型插入点
                else:
                    if p_n in insert_points:
                        raise Exception("Duplicate insert point: " + p_n)
                    insert_points[p_n] = {'name': p_n, 'type': p_t, 'run': r, 'text': '{{' + p + '}}'}

    datas = {
        't1': 'tt1',
        't2': 'tt2',
        't3': 'tt3',
        't4': 'tt4',
        'o1': ['o11', 'o12', 'o13'],
    }

    no_data_points = {}

    for p_n, p_d in insert_points.items():
        if p_n in datas:
            if p_d['type'] == 'text':
                t = p_d['run'].text.replace(p_d['text'], datas[p_n])
                p_d['run'].text = t
            elif p_d['type'] == 'ordered-list':
                print()
                pass
            else:
                raise Exception('solve error')
        else:
            no_data_points[p_n] = p_d

    document.save("data/r3.docx")

    print(no_data_points)


if __name__ == '__main__':
    main()
