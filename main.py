from data_loader import *
from doc_processor import DocumentProcessor
from template_manager import TemplateManager


def main():
    data_loader = JsonDataLoader(file_path='data/json_data.txt')
    TemplateManager.load_info('data/.index')

    tem_path = 'data/template-demo.docx'
    to_dir = "test_data"

    data = data_loader.load_data()
    DocumentProcessor.create_doc_specify_info(to_dir, data, tem_path)


if __name__ == '__main__':
    main()
