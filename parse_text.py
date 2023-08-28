"""
-*- coding: utf-8 -*-
Time     : 2023/8/25
Author   : Hillking
File     : text_parse.py
Function : 提取所有文本内容
"""

from xml.dom.minidom import Element
import os
from parse_numering import NumberParse


class TextParse:

    def __init__(self, numbering_xml_path):
        # 初始化编号分析实例
        self.number_parse = NumberParse()
        if os.path.exists(numbering_xml_path):
            self.number_parse.get_numbering(numbering_xml_path)

        # 段落编号指针
        self.para_pointer = -1
        self.para_info_lst: [[str]] = [] # 内部列表是一个三元组

    def get_content_from_tag_p(self, root_dom: Element, p_text: str = '') -> str:
        # TODO: 修改段落文本解析代码
        '''
        :param root_dom: 节点：不为空节点
        :param p_text: 各阶段的文本
        :return: 返回最终段落的文本
        :func: 解析p标签内部子节点
        '''

        if root_dom is None:
            return p_text

        root_dom_name = root_dom.nodeName
        if root_dom_name == 'w:t':
            dom_text = root_dom.childNodes[0].data
            p_text = p_text + dom_text

        if root_dom_name == 'w:numPr':  # 编号解析
            num_text = self.number_parse.get_parse_number_text(root_dom)
            if '[？]' in num_text:
                print('存在未知编号！===：'+ num_text)

            p_text = p_text + num_text + ' '

        child_doms = root_dom.childNodes
        if len(child_doms) == 0:
            return p_text

        for child_dom in child_doms:
            p_text = self.get_content_from_tag_p(child_dom, p_text)
        return p_text

    def get_content_from_tag_tb(self, root_dom: Element) -> [str, [str]]:
        '''
        :func: 解析表格中的文本并按照某种格式输出
        :param root_dom: tbl表格节点，默认不为空节点
        :return: 表格完整的文本串(同md)， 表格每一行的文本串列表(同md)
        '''

        if root_dom is None:
            raise os.error(2, '未传入表格节点！===：None')
        if root_dom.nodeName != 'w:tbl':
            raise os.error(2, '表格节点类型错误！===：' + root_dom.nodeName)

        tb_row_lst = []
        tb_whole_text = ''
        child_doms = root_dom.childNodes
        for child_dom in child_doms:
            child_dom_name = child_dom.nodeName
            if child_dom_name == 'w:tr':  # 解析每一行的元素
                tb_cel_doms = child_dom.childNodes
                row_text = ''  # 每一行的文本初始化
                for tb_cel_dom in tb_cel_doms:
                    tb_cel_dom_name = tb_cel_dom.nodeName
                    if tb_cel_dom_name == 'w:tc':  # 解析每一行的每一列元素
                        tb_p_doms = tb_cel_dom.getElementsByTagName('w:p')
                        p_text = ''
                        for i, tb_p_dom in enumerate(tb_p_doms):
                            p_text = self.get_content_from_tag_p(tb_p_dom, p_text)
                            if i + 1 != len(tb_p_doms):
                                p_text = p_text + '\n'
                        p_text = ' ' if p_text in ['\n',''] else p_text
                        row_text = row_text + ' | ' + p_text
                row_text = row_text + ' | \n'
                tb_whole_text = tb_whole_text + row_text
                tb_row_lst.append(row_text)

        return tb_whole_text, tb_row_lst

    def get_content_from_tag_sdt(self, root_dom: Element):
        '''
        :func: 从结构体中解析特定文本
        :param root_dom:
        :return:
        '''

        # TODO：补充代码
        pass

    def parse_different_dom(self, root_dom: Element):
        '''
        :param root_dom: 节点
        :return:
        :func: 不同类型的节点判断，并选择不同的节点处理方法.深度优先遍历

        PS: 注意节点层级嵌套问题，p节点内的t文本重复输出，因此只有当p父节点为body时才输出
        '''

        if root_dom is None:
            return

        # Min Dom Levels
        if root_dom.nodeName == 'w:p':
            p_text = self.get_content_from_tag_p(root_dom, '')  # 文本段落处理
            if root_dom.parentNode.nodeName == 'w:body':
                # print(p_text)
                p_text = p_text.strip(' \n')
                if p_text != '':
                    p_struct = []
                    self.para_pointer += 1
                    p_struct.append(self.para_pointer)
                    p_struct.append('P')
                    p_struct.append(p_text)
                    self.para_info_lst.append(p_struct)
        elif root_dom.nodeName == 'w:tbl':
            tb_whole_text, tb_row_lst = self.get_content_from_tag_tb(root_dom)
            # print(tb_whole_text)
            tb_whole_text = tb_whole_text.strip('\n')
            tb_struct = []
            self.para_pointer += 1
            tb_struct.append(self.para_pointer)
            tb_struct.append('TB')
            tb_struct.append(tb_whole_text)
            self.para_info_lst.append(tb_struct)

        elif root_dom.nodeName == 'w:sdt':
            self.get_content_from_tag_sdt(root_dom)

        # TODO: 待扩展的标签：图片、附件
        # elif root_dom.nodeName == '':

        # 子节点循环遍历
        child_doms = root_dom.childNodes
        if len(child_doms) == 0:
            return
        for child_dom in child_doms:
            self.parse_different_dom(child_dom)




if __name__ == '__main__':
    from xml.dom.minidom import parse
    dom_xml = parse('xxx\word\document.xml')
    root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>

    text_parse = TextParse('xxx\word\\numbering.xml')
    text_parse.parse_different_dom(root_dom)  # 根节点初始化
    para_lst = text_parse.para_info_lst
    for para in para_lst:
        print(para)