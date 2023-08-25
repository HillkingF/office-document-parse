"""
-*- coding: utf-8 -*-
Time     : 2023/8/24
Author   : Hillking
File     : parse_numering.py
Function : Parse numbering.xml
"""
import os.path
from xml.dom.minidom import Element, parse
import re

class NumberParse:
    numid_absid_dict: {str: str}  # numid 和 abstractNumId的映射关系
    abstractnum_dict: {str: {str: {str: str}}}  # 每个abstractNumId对应的所有级别及级别的属性
    num_pointer: {str: {str: str}}  # 每种abs中，某一种级别，对应的第几个编号

    def __init__(self):
        self.numid_absid_dict = {}
        self.abstractnum_dict = {}
        self.num_pointer = {}

    def get_num_from_numbering_xml(self, root_dom: Element):
        '''
        :fuc: 解析numbering.xml文件，获取w:numbering节点的直接子节点的编号信息
        :param root_dom:  必须不为空节点
        '''

        # 变量初始化
        num_child_doms = root_dom.childNodes
        for child_dom in num_child_doms:
            child_dom_name = child_dom.nodeName
            if child_dom_name == 'w:num':
                numID = child_dom.getAttribute('w:numId')
                abstractNumId = child_dom.childNodes[0].getAttribute('w:val')
                self.numid_absid_dict[numID] = abstractNumId
            elif child_dom_name == 'w:abstractNum':
                abstractNumId = child_dom.getAttribute('w:abstractNumId')
                level_dict: {str: {str: str}}  # 每个abstractNum的{级别: {属性}}
                level_dict = {}

                lvl_doms = child_dom.childNodes
                for lvl_dom in lvl_doms:
                    level_attr_dict: {str: str}
                    level_attr_dict = {}
                    lvl_dom_name = lvl_dom.nodeName
                    if lvl_dom_name == 'w:lvl':
                        lvl_dom_level = lvl_dom.getAttribute('w:ilvl')
                        level_attr_dict['level'] = lvl_dom_level

                        lvl_child_doms = lvl_dom.childNodes
                        for lvl_child_dom in lvl_child_doms:
                            lvl_child_dom_name = lvl_child_dom.nodeName
                            if lvl_child_dom_name == 'w:start':
                                start = lvl_child_dom.getAttribute('w:val')
                                level_attr_dict['start'] = start
                            elif lvl_child_dom_name == 'w:numFmt':
                                numFmt = lvl_child_dom.getAttribute('w:val')
                                level_attr_dict['numFmt'] = numFmt
                            elif lvl_child_dom_name == 'w:lvlText':
                                lvlText = lvl_child_dom.getAttribute('w:val')
                                # pattern = "(%\d+)"
                                # lvlText = re.sub(pattern, '*', lvlText)
                                level_attr_dict['lvlText'] = lvlText

                        level_dict[lvl_dom_level] = level_attr_dict

                self.abstractnum_dict[abstractNumId] = level_dict

    def get_numbering(self, numbering_xml_path: str = ''):
        if not os.path.exists(numbering_xml_path) and not numbering_xml_path.endswith('.xml'):
            raise os.error(2, "No such file or not a file which like numbering.xml: " + numbering_xml_path)

        num_xml = parse(numbering_xml_path)
        num_dom = num_xml.documentElement

        try:
            self.get_num_from_numbering_xml(num_dom)
        except Exception as e:
            raise os.error(2, "'numbering.xml' parsing error!: " + numbering_xml_path)

    def get_parse_number_text(self, root_dom: Element) -> str:
        '''
        :param root_dom: 具有编号信息的节点。不能为空；必须为 w:numPr
        :return: 返回编号节点对应的真实编号文本
        '''
        if root_dom is None:
            raise os.error(2, '编号节点为空！===：None')

        root_dom_name = root_dom.nodeName
        # print(root_dom_name)
        lvlText_infact = ''
        if root_dom_name == 'w:numPr':  # 编号解析
            num_level = ''
            num_id = ''

            # 获取该节点中的编号级别和编号值
            num_childs = root_dom.childNodes
            for num_child in num_childs:
                num_child_name = num_child.nodeName
                if num_child_name == 'w:ilvl':
                    num_level = num_child.getAttribute('w:val')
                elif num_child_name == 'w:numId':
                    num_id = num_child.getAttribute('w:val')


            # 获取numid和absid的映射关系、编号级别对应的属性
            absnum_id = self.numid_absid_dict.get(num_id)
            if absnum_id is None:
                return ''
            abs_level_attr = self.abstractnum_dict.get(absnum_id).get(num_level)
            if abs_level_attr is None:
                return ''


            start_attr = abs_level_attr.get('start')   # abs_level_attr 属性解析
            numFmt_attr = abs_level_attr.get('numFmt')
            lvlText_attr = abs_level_attr.get('lvlText')


            start_num_infact = ''   # 获取实际编号值，并重置编号指针栏
            if self.num_pointer.get(absnum_id):
                if self.num_pointer.get(absnum_id).get(num_level):
                    # 编号递增
                    start_num_infact = self.num_pointer.get(absnum_id).get(num_level)
                    next_num = int(start_num_infact) + 1
                    self.num_pointer.get(absnum_id)[num_level] = str(next_num)  # 指针为下一个同级别的编号
                else:
                    # 编号重置
                    start_num_infact = start_attr
                    next_num = int(start_num_infact) + 1
                    self.num_pointer.get(absnum_id)[num_level] = str(next_num)
            else:
                start_num_infact = start_attr
                next_num = int(start_num_infact) + 1
                level_value = {}
                level_value[num_level] = str(next_num)
                self.num_pointer[absnum_id] = level_value

            # 解析真实的编号文本
            if numFmt_attr in ['chineseCounting','japaneseCounting']:
                num_dict = {'0': '零', '1': '一', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '七', '8': '八', '9': '九', '10':'十'}
                pattern = "(%\d+)"
                lvlText_infact = re.sub(pattern, num_dict.get(start_num_infact), lvlText_attr)
            elif numFmt_attr in ['decimal']:
                # num_dict = {'0': '0', '1': '1', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '七', '8': '八', '9': '九'}
                pattern = "(%\d+)"
                lvlText_infact = re.sub(pattern, str(start_num_infact), lvlText_attr)  # 模式，替换的文本，编号模板

            else:
                # TODO:补充解析其他类型编号文本的代码
                lvlText_infact = '非中文类型文本暂未解析'
                print(start_attr, numFmt_attr, lvlText_attr, )

            return lvlText_infact
        else:
            raise os.error(2, '编号节点类型错误！===：' + root_dom_name)

