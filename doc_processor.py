from template_analyzer import TemplateAnalyzer


class DocumentProcessor:
    @staticmethod
    def insert_data_to_no_content_point(p_d: dict):
        """校验模板并在预处理阶段插入无内容标签的信息"""
        p_t = p_d['type']
        if p_t in TemplateAnalyzer.insert_point_no_content_types:
            no_content_label = TemplateAnalyzer.registered_labels[p_t]
            no_content_label.insert_data_to_point(p_d, None, TemplateAnalyzer.static_datas)
            return True
        return False

    @staticmethod
    def solve_content_labels(insert_points, datas):
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
    def print_no_data_points(no_data_points):
        """无数据的插入点的默认打印函数"""
        if len(no_data_points) == 0:
            return
        print("无数据对应的内容标签：")
        i = 1
        for point_name, point_data in no_data_points.items():
            print(f"  ({i}) 标签名为'{point_name}'、类型为'{point_data['type']}'的内容标签'{point_data['text']}'无法匹配到数据，原文：{point_data['run'].text}")
