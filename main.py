import sys

from data_loader import *
from doc_processor import DocumentProcessor
# from template_manager import TemplateManager


def write_result_to_file(file_path, result: [str, int, bool]):
    with open(file_path, 'w') as f:
        f.write(str(result))


def main():
    sys.stdout = None
    sys.stderr = None

    args = sys.argv[1:]
    help_str = '参数格式为：模板路径，生成目录，文档名称，数据文件路径，结果输出文件路径。'
    if len(args) == 1 and args[0] == '-h':
        print(help_str)
        return False
    elif len(args) == 5:
        t_p, d_p, d_n, j_p, r_p = args
    else:
        print('没有提供创建文档参数或参数错误。 ' + help_str)
        return False

    # data_loader = JsonDataLoader(file_path='data/json_data.txt')
    data_loader = JsonDataLoader(file_path=j_p)
    # TemplateManager.load_info('data/.index')

    # tem_path = 'data/template-demo.docx'
    tem_path = t_p
    # to_dir = "test_data"
    to_dir = d_p

    # 通过魔法数据传递指定的模板参数
    data = data_loader.load_data()
    if data is not None:
        result = DocumentProcessor.create_doc_specify_info(to_dir, data, tem_path=tem_path, doc_name=d_n)
        write_result_to_file(r_p, int(result))
        return result
    else:
        print('数据为空，无法创建文档。')
        write_result_to_file(r_p, int(False))
        return False


if __name__ == '__main__':
    main()
