"""
-*- coding: utf-8 -*-
Time     : 2023/9/1
Author   : Hillking
File     : merge_mark2txt.py
Function : 将文档标注和文档解析合并成索引+内容的格式
"""
import os
import re


class MarkStack:
    def __init__(self):
        self.stack = []
        self.stack_max_index = -1      # 记录栈对象从创建后的最大位置索引，当索引不足需要新增元素时，使用append,否则直接array[i]
        self.stack_top_pointer = 0   # 栈顶指针始终指向当前轮遍历中最后一个有值的位置
        self.title = ''

    def pattern_level(self, mark_txt_row):
        '''
        : stack_level: 标题级别，【章节】【标题】【xx级】
        : stack_num: 章节编号
        : stack_content: 具体的标题内容
        '''
        # 匹配章节级别
        pattern = '【标题】|【章节】|【[0-9]+级】'
        stack_text = re.findall(pattern, mark_txt_row)
        if len(stack_text) == 0:
            return '', None, None
        elif len(stack_text) > 1:
            raise (2, '标注错误！' + mark_txt_row)
        else:
            stack_level = stack_text[0]

        stack_num_part = re.findall('\d+', stack_level)
        if len(stack_num_part) == 0 and stack_level == '【章节】':
            stack_num = 0
        elif len(stack_num_part) == 0 and stack_level == '【标题】':
            stack_num = -1
        elif len(stack_num_part) == 1:
            stack_num = int(stack_num_part[0])
        else:
            raise (2, '编号解析错误！' + mark_txt_row)

        stack_content = mark_txt_row.replace(stack_level, '').strip('\n\t')
        return stack_level, stack_num, stack_content

    def push(self, mark_txt_row):
        # 匹配章节级别
        stack_level, stack_num, stack_content = self.pattern_level(mark_txt_row)

        if stack_num == -1:
            pass  # 说明是标题
        elif stack_num == 0:
            # 判断级别是否为【章节】，如果是直接重置栈和指针
            self.stack = []
            self.stack_top_pointer = 0
            self.stack.append(stack_content)
            self.stack_max_index = 0
            # print(len(self.stack), self.stack_top_pointer)

        elif stack_num > 0:  # 如果是小节
            if self.stack_max_index < stack_num:    # 判断是否超过最大深度
                while self.stack_max_index < stack_num:  # 如果不连贯则一直填入空字符串
                    self.stack.append('')
                    self.stack_top_pointer += 1
                    self.stack_max_index += 1
                if self.stack_max_index == stack_num:    # 直到连贯再更新
                    self.stack[self.stack_top_pointer] = stack_content
            else:  # 如果没有超过最大深度
                while self.stack_top_pointer < stack_num: # 如果不连贯则一直填入空字符串
                    self.stack_top_pointer += 1
                    self.stack[self.stack_top_pointer] = ''
                if self.stack_top_pointer == stack_num:
                    # print(len(self.stack), self.stack_top_pointer)
                    self.stack[self.stack_top_pointer] = stack_content

                if self.stack_top_pointer > stack_num:  # 说明要出栈了
                    self.stack[stack_num] = stack_content
                    self.stack_top_pointer = stack_num

    def title_index_text(self):
        # 栈底到栈顶的文字，使用\n分隔，用于创建索引
        index = 0
        title_index = ''
        if len(self.stack) == 0:
            title_index = ''
            return title_index

        while index <= self.stack_top_pointer:
            title_index = title_index + '#' + self.stack[index]
            index += 1
        title_index = title_index.strip('#')
        return title_index


class MergeParseAndMark:
    def __init__(self, para_lst, mark_file_path):
        '''
        :param para_lst: docx解析后的三元组
        :param mark_file_path:  标注文档路径

        :var self.mark_stack: 每个文档对应的栈
        :var self.mark_lines: 标注文档的所有行
        :var self.mark_lines_pointer: 标注行指针
        :var self.para_lst: 解析的docx文档三元组
        '''
        self.mark_stack = MarkStack()    # 创建一个空的栈

        if not os.path.exists(mark_file_path):
            raise (2, '标注文档路径不存在：' + mark_file_path)
        with open(mark_file_path, 'r') as f:   # , encoding='utf-8'
            self.mark_lines = f.readlines()

        self.mark_lines_pointer = 0
        self.mark_last_line_attr = ''

        self.para_lst = para_lst

        self.title = ''

        # 新三元组
        self.index_triplet = []
        # self.triplet_1 = ''
        self.triplet_2 = []
        self.triplet_3 = []




    def build_index_content(self):
        '''
        :var each_triplet: 文本解析三元组
        :var para_type_lst: 每个三元组中的类型列表
        :var para_content_lst: 每个三元组中的内容列表
        :return:
        '''
        # 遍历三元组的每一行，获取、类型、内容和索引

        # 1.遍历三元组的每一行
        for i, each_triplet in enumerate(self.para_lst):
            para_type_lst = each_triplet[1]
            para_content_lst = each_triplet[2]


            # 2、判断每个三元组的内容元素数量
            len_content = len(para_content_lst)

            if self.mark_lines_pointer < len(self.mark_lines):
                mark_row = self.mark_lines[self.mark_lines_pointer]
            else:
                mark_row = ''
            stack_level, stack_num, stack_content = self.mark_stack.pattern_level(mark_row)

            # 3.标题更新
            if len_content == 1 and para_type_lst[0] == 'p' and stack_content is not None and para_content_lst[0].strip('\n\t').replace(' ','') == stack_content.replace(' ',''):
                # print('是章节', stack_level, stack_num, stack_content, para_content_lst[0].strip('\n\t').replace(' ',''))

                if stack_num == -1:  # 说明是标题
                    self.title = stack_content
                else:
                    if self.mark_last_line_attr == '文本段':
                        # 构建新的三元组行
                        triplet = []
                        triplet.append(self.title)
                        triplet.append(self.mark_stack.title_index_text())
                        triplet.append(self.triplet_2)
                        triplet.append(self.triplet_3)
                        self.index_triplet.append(triplet)

                        # 更新三元组元素
                        # self.triplet_1 = ''
                        self.triplet_2 = []
                        self.triplet_3 = []
                    # 更新栈元素
                    self.mark_stack.push(mark_row)
                self.mark_lines_pointer += 1
                self.mark_last_line_attr = '章节'
            else:
                # 合并内容
                self.triplet_2.extend(para_type_lst)
                self.triplet_3.extend(para_content_lst)
                self.mark_last_line_attr = '文本段'

        # 文本段落收尾
        if self.triplet_2 != []:
            # 构建新的三元组行
            triplet = []
            triplet.append(self.title)
            triplet.append(self.mark_stack.title_index_text())
            triplet.append(self.triplet_2)
            triplet.append(self.triplet_3)
            self.index_triplet.append(triplet)
            # 更新三元组元素
            # self.triplet_1 = ''
            self.triplet_2 = []
            self.triplet_3 = []


    def get_index_triplet(self):
        return self.index_triplet



if __name__ == '__main__':
    root = 'D:\Project\Dataset\排版-完整版财务知识'
    pre_file_dir_name = '原版所有文档'
    unzip_file_path = os.path.join(root, '压缩包解压目录')

    result_dir = os.path.join(root, '语句解析结果')
    docxs_dir = unzip_file_path



    from xml.dom.minidom import parse
    from parse_text import TextParse
    dir = '2 贝壳集团核算月结管理制度 - 副本'   #'2 贝壳集团核算月结管理制度'
    docu_xml = os.path.join(docxs_dir, dir, 'word', 'document.xml')
    numb_xml = os.path.join(docxs_dir, dir, 'word', 'numbering.xml')
    docu_xml_rels = os.path.join(docxs_dir, dir, 'word')
    final_annex_dir = os.path.join(result_dir, dir)
    dom_xml = parse(docu_xml)
    root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>
    text_parse = TextParse(numb_xml, docu_xml_rels, final_annex_dir)
    text_parse.parse_different_dom(root_dom)  # 根节点初始化
    para_lst = text_parse.para_info_lst


    mark_file_path = 'C:\\Users\\fengwenni001\Desktop\BAC文档标注说明\\2 贝壳集团核算月结管理制度.txt'
    merge = MergeParseAndMark(para_lst, mark_file_path)
    merge.build_index_content()
    index_and_content = merge.get_index_triplet()
    # for x in index_and_content:
    #     print(x)


    # 序列化保存该四元组列表
    import pickle
    with open('triplets.txt', 'wb') as f1:
        pickle.dump(index_and_content, f1)
