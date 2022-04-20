from docx import Document

from helper.os_helper import *
from magic_data import MagicData
from template_analyzer import TemplateAnalyzer
from template_manager import TemplateManager


class DocumentProcessor:
    """文档处理器，负责修改文档中的内容"""
    @staticmethod
    def _insert_no_content_data(p_d: dict):
        """校验模板并在预处理阶段插入无内容标签的信息"""
        p_t = p_d['type']
        if p_t in TemplateAnalyzer.insert_point_no_content_types:
            no_content_label = TemplateAnalyzer.registered_labels[p_t]
            no_content_label.insert_data_to_point(p_d, None, TemplateAnalyzer.static_datas)
            return True
        return False

    @staticmethod
    def _solve_content_labels(insert_points, datas):
        """处理有内容类型插入点，检查并插入数据"""
        no_data_points = {}
        for point_name, point_data in insert_points.items():
            if point_name in datas:
                data = datas[point_name]
                label = TemplateAnalyzer.registered_labels[point_data['type']]
                if not label.check_data_type(data):
                    print(f"插入内容类型 {point_data['type']} 不能匹配数据 {type(data)}。内容标签为 {point_data['text']}，原文为 {point_data['run'].text}")
                    continue
                label.insert_data_to_point(point_data, data, TemplateAnalyzer.static_datas)
            else:
                no_data_points[point_name] = point_data
        return no_data_points

    @staticmethod
    def _print_no_data_points(no_data_points):
        """无数据的插入点的默认打印函数"""
        if len(no_data_points) == 0:
            return
        print("无数据对应的内容标签：")
        i = 1
        for point_name, point_data in no_data_points.items():
            print(f"  ({i}) 标签名为'{point_name}'、类型为'{point_data['type']}'的内容标签'{point_data['text']}'无法匹配到数据，原文：{point_data['run'].text}")

    @staticmethod
    def _save_document(save_dir: str, document: Document, magic_data: MagicData):
        save_path = f'{save_dir}/{magic_data.get_doc_name()}'
        make_sure_path(os.path.dirname(save_path))
        document.save(save_path)

    @staticmethod
    def create_doc(save_dir: str, datas: dict):
        """
        根据模板创建word文档，参数使用魔法数据指定

        :param save_dir: 保存文件路径的基础路径
        :param datas: 待插入模板的数据
        :return: 是否创建成功
        """
        if not MagicData.check_all_magic_data(datas):
            print("创建文件的魔法数据没有设置完整")
            return False

        make_sure_path(save_dir)

        # 更新模板分析器的标签信息
        # TemplateAnalyzer.update_labels_info()

        # 注册每次模板生成过程的静态插入数据
        TemplateAnalyzer.register_static_datas()

        # 解析魔法数据
        magic_data = MagicData(datas)

        # 模板检查与预处理
        tem_path = TemplateManager.get_template_path(magic_data.get_tem_name())
        check_result = TemplateAnalyzer.check_template(tem_path, DocumentProcessor._insert_no_content_data)
        TemplateAnalyzer.print_check_info(check_result, show_detail=True)
        # 模板校验失败直接退出
        if check_result['code'].is_error():
            return False
        insert_points = check_result['data']['insert_points']
        document = check_result['data']['document']

        # 处理有内容类型插入点，检查并插入数据，返回没有对应数据的插入点
        no_data_points = DocumentProcessor._solve_content_labels(insert_points, datas)
        # 打印没有数据的插入点
        DocumentProcessor._print_no_data_points(no_data_points)

        # 保存文件，文件若存在，则会覆盖文件
        DocumentProcessor._save_document(save_dir, document, magic_data)

        return True

    @staticmethod
    def create_doc_specify_info(to_dir, data, tem_path=None, doc_name=None):
        """
        根据模板创建word文档，可指定参数

        :param to_dir: 存储的基础目标路径
        :param data: 文档的插入数据
        :param tem_path: 模板路径，若为空则要求数据中已经设置模板名称的魔法数据
        :param doc_name: 文档名称，若为空则要求数据中已经设置文档名称的魔法数据
        :return: 创建是否成功
        """
        if tem_path is not None:
            # 指定不在模板库的模板
            tem_n = TemplateManager.add_template(tem_path)
            if tem_n is None or not MagicData.set_tem_name(data, tem_n, is_cover=True):
                print(f"模板文件路径错误：{tem_path}")
                return False

        # 手动指定生成文件名
        if doc_name is not None:
            if not MagicData.set_doc_name(data, doc_name, is_cover=True):
                print("文档名设置错误")
                return False

        return DocumentProcessor.create_doc(to_dir, data)
