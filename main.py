from helper.os_helper import *
from data_loader import StaticDataLoader
from template_analyzer import TemplateAnalyzer
from doc_processor import DocumentProcessor


def match(file_path: str, save_path: str, datas: dict):
    # 更新模板分析器的标签信息
    # TemplateAnalyzer.update_labels_info()

    # 注册每次模板生成过程的静态插入数据
    TemplateAnalyzer.register_static_datas()

    # 模板检查与预处理
    check_result = TemplateAnalyzer.check_template(file_path, DocumentProcessor.insert_data_to_no_content_point)
    TemplateAnalyzer.print_check_info(check_result, True)
    # 模板校验失败直接退出
    if check_result['code'].is_error():
        return
    insert_points = check_result['data']['insert_points']
    document = check_result['data']['document']

    # 处理有内容类型插入点，检查并插入数据，返回没有对应数据的插入点
    no_data_points = DocumentProcessor.solve_content_labels(insert_points, datas)
    # 打印没有数据的插入点
    DocumentProcessor.print_no_data_points(no_data_points)

    # 保存文件
    document.save(save_path)


def main():
    datas = [
        {
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
        },
    ]

    data_loader = StaticDataLoader(datas)

    tem_path = "data/template-demo.docx"
    to_dir = "test_data"
    word_name = "word.docx"
    make_sure_path(to_dir)

    match(tem_path, f"{to_dir}/{word_name}", data_loader.load_data())


if __name__ == '__main__':
    main()
