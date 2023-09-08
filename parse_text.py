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
from parse_image import ImageParse
import pandas as pd

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



        # annex 关系
        self.annex_example = AnnexParse(annex_dir)
        self.annex_relation = self.annex_example.get_annex_relation()

        # 附件转储根目录
        self.annex_final_dir = final_annex_dir
        if self.annex_relation != {} and not os.path.exists(self.annex_final_dir):
            os.mkdir(self.annex_final_dir)
            self.annex_example.copy_annex(self.annex_final_dir)

        # 图像分析
        self.image_parse = ImageParse()


        # 段落编号指针
        self.para_pointer = -1
        self.para_info_lst: [[str]] = [] # 内部列表是一个三元组




    def get_content_from_tag_p(self, root_dom: Element, p_type: [str]=[], p_text: [str]=[] ) -> [[str], [str]]:
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
                    attr_name = attr_rela.get('Target').split('/')[-1]

                    if attr_type.split('/')[-1] == 'image':
                        p_type.append(TYPE_IMAGE)
                        # p_text.append(os.path.join(self.annex_final_dir, attr_name))
                        p_text.append(attr_name)   # 附件元组文本内容
                    else:
                        p_type.append(TYPE_FILE)
                        # p_text.append(os.path.join(self.annex_final_dir, attr_name))
                        p_text.append(attr_name)
                    return p_type, p_text
        elif root_dom_name == 'w:drawing':
            file_type, file_name = self.image_parse.get_image_path(root_dom, self.annex_relation)
            image_type, serialized_image = self.image_parse.image_serialize(os.path.join(self.annex_final_dir, file_name))
            # 保存序列化后的图像
            # with open('serialized_image.txt', 'wb') as serialized_image_file:
            #     serialized_image_file.write(serialized_image.encode())
            image_lst = []
            image_lst.append(file_name)
            image_lst.append(serialized_image)
            p_type.append(file_type)
            p_text.append(image_lst)
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
        :return: 表格结构体，字典
        '''

        if root_dom is None:
            raise os.error(2, '未传入表格节点！===：None')
        if root_dom.nodeName != 'w:tbl':
            raise os.error(2, '表格节点类型错误！===：' + root_dom.nodeName)

        # table struct init.
        tb_struct = {}
        tb_struct['type'] = 'TB'        # mark tb type
        tb_struct['name'] = ''          # tb name, default ''
        tb_struct['row_cnts'] = 0       # tb row counts, init 0
        tb_struct['col_cnts'] = 0       # tb column counts, init 0
        tb_struct['header_num'] = 0     # tb header number, default the first row
        tb_struct['column_num'] = -1    # tb column number, default no, mark with -1
        tb_struct['content'] = []       # tb main content, type-list. each element indicate one row and its length is the number of columns of tb . eg: [[cel1, cel2, ...],[2th row],[3th row],[...],[...],[...]].   tb_cel: {'text':'', 'vmerge':'', 'hmerge': ''}

        # 索引专用文本list
        index_text_list = []

        # 获取列数
        tmp_tblGrid_cnts = 0
        child_doms = root_dom.childNodes
        for child_dom in child_doms:
            child_dom_name = child_dom.nodeName
            if child_dom_name == 'w:tblGrid':
                tmp_tblGrid_cnts += 1
                col_cnts = 0
                gridCol_doms = child_dom.childNodes
                for gridCol_dom in gridCol_doms:
                    if gridCol_dom.nodeName == 'w:gridCol':
                        col_cnts += 1
                tb_struct['col_cnts'] = col_cnts
            elif child_dom_name == 'w:tr':
                # 遍历每一行
                tb_struct['row_cnts'] = tb_struct.get('row_cnts') + 1  # 行号+1
                tr_child_doms = child_dom.childNodes

                index_row_lst = []
                tr_cel_lst = []
                for tr_child_dom in tr_child_doms:
                    if tr_child_dom.nodeName == 'w:tc':
                        tc_child_doms = tr_child_dom.childNodes
                        tcpr_nums = 0
                        p_text = []
                        p_type = []
                        cel_text = ''
                        gridSpan_nums = 1
                        vMerge_dom_format = 'single'
                        for tc_child_dom in tc_child_doms:  # tcPr、p
                            if tc_child_dom.nodeName == 'w:tcPr':
                                tcpr_nums += 1
                                tcPr_child_doms = tc_child_dom.childNodes
                                # == 单元格格式 ==
                                for tcpr_child_dom in tcPr_child_doms:
                                    if tcpr_child_dom.nodeName == 'w:gridSpan':
                                        gridSpan_nums = int(tcpr_child_dom.getAttribute('w:val'))
                                    elif tcpr_child_dom.nodeName == 'w:vMerge':
                                        vMerge_dom_format = tcpr_child_dom.getAttribute('w:val')

                            elif tc_child_dom.nodeName == 'w:p':
                                if len(p_text) != 0:
                                    p_text[0] = p_text[0] + '\n'
                                p_type, p_text = self.get_content_from_tag_p(tc_child_dom, p_type, p_text)


                        if len(p_text) != 0:  # 单元格不为空。由于默认单元格中只有文字，因此p_text中只有一个元素
                            cel_text = p_text[0].strip('\n')
                        else:
                            cel_text = ''

                        # 单元格结构体
                        for c in range(1, gridSpan_nums + 1):  # 水平合并单元格构造
                            cel_struct = {}
                            if c == 1:
                                cel_struct['text'] = cel_text
                                cel_struct['v_format'] = vMerge_dom_format
                                cel_struct['h_format'] = 'single' if gridSpan_nums == 1 else 'restart'
                                index_row_lst.append(cel_text)
                            else:
                                cel_struct['text'] = ''
                                cel_struct['v_format'] = vMerge_dom_format
                                cel_struct['h_format'] = 'continue'
                                index_row_lst.append('')
                            tr_cel_lst.append(cel_struct)
                # 文本list、结构体lst
                tb_struct.get('content').append(tr_cel_lst)
                index_text_list.append(index_row_lst)

        if tmp_tblGrid_cnts != 1:  # 列数标签验证
            raise (2, 'tbl表格的列数标签组<w:tblGrid>数量不为1！')

        return index_text_list, tb_struct

    def get_content_from_tag_sdt(self, root_dom: Element):
        '''
        :func: 从结构体中解析特定文本
        :param root_dom:
        :return:
        '''
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
            index_text_list, tb_struct = self.get_content_from_tag_tb(root_dom)   # 文本lst和结构体list
            # 将行列元素列表tb_row_lst_pd转换成dataframe格式
            tb_row_pd = pd.DataFrame(index_text_list, columns=None, index=None, dtype=str)
            tb_row_pd_str = tb_row_pd.to_string(header=False,index=False)

            # # 字典和json互换
            # import json
            # dic2json = json.dumps(tb_struct, ensure_ascii = False)
            # json2dict = json.loads(dic2json)

            tb_type = ['TB']
            tb_text = [tb_row_pd_str, tb_struct]

            p_struct = []
            self.para_pointer += 1
            p_struct.append(self.para_pointer)
            p_struct.append(tb_type)
            p_struct.append([tb_text])  # 表格内容结构体：包含df序列化后的表格文本和表格完整结构体
            self.para_info_lst.append(p_struct)

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