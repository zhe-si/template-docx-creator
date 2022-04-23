import sys

from data_loader import *
from doc_processor import DocumentProcessor
# from template_manager import TemplateManager


def main():
    args = sys.argv[1:]
    help_str = '参数格式为：模板路径，生成目录，文档名称，数据文件路径。'
    if len(args) == 1 and args[0] == '-h':
        print(help_str)
        return
    elif len(args) == 4:
        t_p, d_p, d_n, j_p = args
    else:
        print('没有提供创建文档参数或参数错误。 ' + help_str)
        return

    # data_loader = JsonDataLoader(file_path='data/json_data.txt')
    data_loader = JsonDataLoader(file_path=j_p)
    # TemplateManager.load_info('data/.index')

    # tem_path = 'data/template-demo.docx'
    tem_path = t_p
    # to_dir = "test_data"
    to_dir = d_p

    # 通过魔法数据传递指定的模板参数
    data = data_loader.load_data()
    DocumentProcessor.create_doc_specify_info(to_dir, data, tem_path=tem_path, doc_name=d_n)


if __name__ == '__main__':
    main()
