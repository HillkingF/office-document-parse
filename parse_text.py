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
from parse_annex import AnnexParse

# 解析内容类型常量
TYPE_TEXT = 'p'
TYPE_FILE = 'F'
TYPE_IMAGE = 'I'


class TextParse:

    def __init__(self, numbering_xml_path, annex_dir, final_annex_dir):
        # 初始化编号分析实例
        self.number_parse = NumberParse()
        if os.path.exists(numbering_xml_path):
            self.number_parse.get_numbering(numbering_xml_path)

        # 附件路径

        # annex 关系
        self.annex_example = AnnexParse(annex_dir)
        self.annex_relation = self.annex_example.get_annex_relation()

        # 附件转储根目录
        self.annex_final_dir = final_annex_dir
        if self.annex_relation != {} and not os.path.exists(self.annex_final_dir):
            os.mkdir(self.annex_final_dir)
            self.annex_example.copy_annex(self.annex_final_dir)


        # 段落编号指针
        self.para_pointer = -1
        self.para_info_lst: [[str]] = [] # 内部列表是一个三元组




    def get_content_from_tag_p(self, root_dom: Element, p_type: [str]=[], p_text: [str]=[] ) -> [[str], [str]]:
        # TODO: 修改段落文本解析代码
        '''
        :param root_dom: 节点：不为空节点
        :param p_text: 各阶段的文本
        :return: 返回最终段落的文本
        :func: 解析p标签内部子节点
        '''

        if root_dom is None:
            return p_type, p_text

        root_dom_name = root_dom.nodeName

        if root_dom_name == 'w:t':
            dom_text = root_dom.childNodes[0].data
            if dom_text == '' or dom_text is None:
                return p_type, p_text

            if len(p_type)==0:  # 新增
                p_type.append(TYPE_TEXT)
                p_text.append(dom_text)
            elif len(p_type) != 0 and p_type[-1] == TYPE_TEXT:  # 同文本类型迭代
                p_text[-1] = p_text[-1] + dom_text
            elif len(p_type) != 0 and p_type[-1] != TYPE_TEXT:  # 异文件类型新增
                p_type.append(TYPE_TEXT)
                p_text.append(dom_text)

            return p_type, p_text

        elif root_dom_name == 'w:numPr':  # 编号解析
            num_text = self.number_parse.get_parse_number_text(root_dom)
            if num_text == '' or num_text is None:
                return p_type, p_text

            if '[？]' in num_text:
                print('存在未知编号！===：'+ num_text)

            if len(p_type)==0:  # 新增
                p_type.append(TYPE_TEXT)
                p_text.append(num_text)
            elif len(p_type) != 0 and p_type[-1] == TYPE_TEXT:  # 同文本类型迭代
                p_text[-1] = p_text[-1] + num_text
            elif len(p_type) != 0 and p_type[-1] != TYPE_TEXT:  # 异文件类型新增
                p_type.append(TYPE_TEXT)
                p_text.append(num_text)
            return p_type, p_text

        elif root_dom_name == 'o:OLEObject':
            attr_id = root_dom.getAttribute('r:id')
            if attr_id == '' or attr_id is None:
                print('附件id为空！')
                return p_type, p_text
            else:
                if not self.annex_relation.get(attr_id):
                    print('附件id对应的关系为空！', attr_id)
                    return p_type, p_text
                else:
                    attr_rela = self.annex_relation.get(attr_id)
                    # analyse attr
                    attr_type = attr_rela.get('Type')
                    attr_name = attr_rela.get('Target')

                    if attr_type.split('/')[-1] == 'image':
                        p_type.append(TYPE_IMAGE)
                        # p_text.append(os.path.join(self.annex_final_dir, attr_name))
                        p_text.append(attr_name)   # 附件元组文本内容
                    else:
                        p_type.append(TYPE_FILE)
                        # p_text.append(os.path.join(self.annex_final_dir, attr_name))
                        p_text.append(attr_name)
                    return p_type, p_text

        child_doms = root_dom.childNodes
        if len(child_doms) == 0:
            return p_type, p_text

        for child_dom in child_doms:
            p_type, p_text = self.get_content_from_tag_p(child_dom, p_type, p_text)
        return p_type, p_text

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
                        p_text = []
                        p_type = []
                        for i, tb_p_dom in enumerate(tb_p_doms):
                            p_type, p_text = self.get_content_from_tag_p(tb_p_dom, p_type, p_text)

                            # if i + 1 != len(tb_p_doms):
                            #     if len(p_text) == 0:
                            #         p_text.append(' ')
                            #         p_type.append(TYPE_TEXT)

                        # p_text[-1] = ' ' if p_text[-1] == '' else p_text
                        if len(p_text) == 0:
                            row_text = row_text + ' | ' + ' '  # 这里默认为表格中只有文本，不存在其他附件！！！
                        else:
                            row_text = row_text + ' | ' + p_text[-1]  # 这里默认为表格中只有文本，不存在其他附件！！！
                        # x = p_text[-1]
                        # row_text = row_text + ' | ' + x   # 这里默认为表格中只有文本，不存在其他附件！！！
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
            p_type, p_text = self.get_content_from_tag_p(root_dom, [] , [])  # 文本段落处理 ,这里默认该段落初始类型为文本，
            if root_dom.parentNode.nodeName == 'w:body':
                # print(p_text)
                if len(p_type) != 0:
                    p_struct = []
                    self.para_pointer += 1
                    p_struct.append(self.para_pointer)
                    p_struct.append(p_type)
                    p_struct.append(p_text)
                    self.para_info_lst.append(p_struct)
        elif root_dom.nodeName == 'w:tbl':
            tb_whole_text, tb_row_lst = self.get_content_from_tag_tb(root_dom)
            # print(tb_whole_text)
            tb_whole_text = tb_whole_text.strip('\n')
            tb_type = ['TB']
            tb_text = [tb_whole_text]

            tb_struct = []
            self.para_pointer += 1
            tb_struct.append(self.para_pointer)
            tb_struct.append(tb_type)
            tb_struct.append(tb_text)
            self.para_info_lst.append(tb_struct)

        elif root_dom.nodeName == 'w:sdt':
            self.get_content_from_tag_sdt(root_dom)


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