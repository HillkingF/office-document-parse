"""
-*- coding: utf-8 -*-
Time     : 2023/8/23
Author   : Hillking
File     : main.py
Function : docx文档解析

https://docs.python.org/zh-cn/3/library/xml.dom.minidom.html
"""

import os
from xml.dom.minidom import parse, Element
from parse_text import TextParse
import trans_docx2xml


# 全局变量库
dom_rely_dict: {str:list} = {}

# 节点库
def dom_rep():
    # body节点的子节点库
    body_child = ['w:p','w:tbl','w:sdt','w:bookmarkStart','w:bookmarkEnd','w:sectPr']
    dom_rely_dict = {}
    return body_child, dom_rely_dict


def set_recurrent_dom(root_dom: Element) -> None:
    """
    循环遍历节点，返回每种节点的所有子节点.
    其中：root_dom将作为父级节点来获取其直接子节点
    默认该节点不为空
    :return: [dom_lst]
    """

    # 获取节点的名称
    root_dom_name = root_dom.nodeName
    # 判断节点是否在节点字典中
    if root_dom_name not in dom_rely_dict.keys():
        dom_rely_dict[root_dom_name] = []

    # 获取该节点的直接子节点，并依次遍历
    child_doms = root_dom.childNodes
    if len(child_doms) == 0:
        return
    else:
        for child_dom in child_doms:
            child_dom_name = child_dom.nodeName
            if child_dom_name not in dom_rely_dict.get(root_dom_name):
                dom_rely_dict[root_dom_name].append(child_dom_name)

            set_recurrent_dom(child_dom)
        return

def get_recurrent_dom(root_dom: Element) -> None:
    """
    循环遍历节点，返回每个节点的所有直接子节点列表.
    其中：root_dom将作为父级节点来获取其直接子节点
    默认该节点不为空
    :return: [dom_lst]
    """

    # 获取节点的名称
    root_dom_name = root_dom.nodeName

    # 获取该节点的直接子节点，并依次遍历
    child_doms = root_dom.childNodes

    if len(child_doms) == 0:
        return
    else:
        for child_dom in child_doms:
            child_dom_name = child_dom.nodeName
            if child_dom_name not in dom_rely_dict.get(root_dom_name):
                dom_rely_dict[root_dom_name].append(child_dom_name)

            set_recurrent_dom(child_dom)
        return

# 遍历所有的xml文档，將所有的节点依赖关系保存到字典及文件中
def parse_xml_first_tag(xml_dir:str='') -> None:
    # 定义方法内的全局变量
    tag_lst = []
    main_doms: [Element]  # body直接子节点

    # 验证docx的document.xml中的所有标签是否有不同库但同名的情况
    for root, dirs, files in os.walk(xml_dir):
        dirs = sorted(dirs)
        for dir in dirs:
            document_file = xml_dir + dir + '\\word\document.xml'
            if os.path.exists(document_file):
                # 解析XML文件，获取文件根节点
                dom_xml = parse(document_file)
                root_dom = dom_xml.documentElement   # 获取文档唯一的根节点 <w:document>
                root_dom_name = root_dom.nodeName
                # 判断根节点是否唯一
                if root_dom_name != 'w:document':
                    print(dir)
                    print("根节点不为w:document")
                    return


                # 获取根节点的唯一子节点：w:body
                body_doms = root_dom.childNodes
                if len(body_doms) > 1 or len(body_doms) == 0:
                    print('根目录的直接子节点数量为' + str(len(body_doms)) + '个！')
                    return
                for x in body_doms:
                    if x.nodeName != 'w:body':
                        print(dir)
                        print('根目录后的节点不是w:body')
                        return
                body_dom = body_doms[0]

                # 获取body节点的所有直接子节点, 直接子节点名称,存入节点库
                main_doms = body_dom.childNodes
                for i, main_dom in enumerate(main_doms):
                    main_dom_name = main_dom.nodeName
                    if main_dom_name not in tag_lst:
                        tag_lst.append(main_dom_name)


                # 遍历所有的子节点
                set_recurrent_dom(body_dom)

def get_text_from_dom(xml_path:str=''):
    dom_xml = parse(xml_path)
    root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>
    root_dom_name = root_dom.nodeName

    # get_content_from_tag_p(root_dom)


if __name__ == '__main__':
    # Args：以下三个目录是必须有的，分别为各种文件的根目录， 所有文档的目录名字，解压缩后的所有xml文件路径
    root = 'D:\Project\Dataset\排版-完整版财务知识'
    pre_file_dir_name = '原版所有文档'
    unzip_file_path = os.path.join(root, '压缩包解压目录')

    # 1、文件zip->unzip
    trans_docx2xml.docx_zip_unzip(root, pre_file_dir_name, unzip_file_path)

    # 2.解析所有的标签，获取文本内容
    result_dir = os.path.join(root, '语句解析结果')
    docxs_dir = unzip_file_path

    dirs = os.listdir(docxs_dir)
    for dir in dirs:
            print('【', dir + '】')

            docu_xml = os.path.join(docxs_dir, dir, 'word', 'document.xml')
            numb_xml = os.path.join(docxs_dir, dir, 'word', 'numbering.xml')
            docu_xml_rels = os.path.join(docxs_dir, dir, 'word')
            final_annex_dir = os.path.join(result_dir, dir)

            save_path = os.path.join(result_dir, dir) + '.txt'

            dom_xml = parse(docu_xml)
            root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>

            text_parse = TextParse(numb_xml, docu_xml_rels, final_annex_dir)
            text_parse.parse_different_dom(root_dom)  # 根节点初始化
            para_lst = text_parse.para_info_lst

            with open(save_path, 'w', encoding='utf-8') as fw:
                for para in para_lst:
                    fw.write(str(para) + '\n')
