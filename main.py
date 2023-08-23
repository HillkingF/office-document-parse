"""
-*- coding: utf-8 -*-
Time     : 2023/8/23
Author   : Hillking
File     : main.py
Function : docx文档解析
"""
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.dom.minidom import parse   ##
import xml.dom.minidom

# https://docs.python.org/zh-cn/3/library/xml.dom.minidom.html


# 节点库
def dom_rep():
    # body节点的子节点库
    body_child = ['w:p','w:tbl','w:sdt','w:bookmarkStart','w:bookmarkEnd','w:sectPr']
    dom_rely_dict = {}
    return body_child, dom_rely_dict


def docx_zip_unzip(root: str='', pre_files_dir:str='') -> None:
    # 遍历所有的子文档，进行压缩，暂存后解压
    predir = root + '\\' + pre_files_dir + '\\'
    zipdir = root + '\压缩包\\'
    unzipdir = root + '\压缩包解压目录\\'

    if not os.path.exists(zipdir):
        os.mkdir(zipdir)
    if not os.path.exists(unzipdir):
        os.mkdir(unzipdir)

    for _, _, files in os.walk(predir):
        for file in files:
            file_name = file.split('.')[0]
            old_file_path = predir + file
            new_file_path = zipdir + file.split('.')[0] + '.zip'
            shutil.copy(old_file_path, new_file_path)  # 将所有的docx文档转换成zip格式的文件
            with zipfile.ZipFile(new_file_path, 'r') as zip_ref:
                # 解压缩文件到指定目录
                zip_ref.extractall(unzipdir + file_name)


dom_rely_dict : {str:list}= {}
def recurrent_dom(root_dom):
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

            recurrent_dom(child_dom)
        return






# 遍历所有的xml文档，分析同级标签
def parse_xml_first_tag(xml_dir:str=''):
    # 定义方法内的全局变量
    tag_lst = []
    tag_lst_only = []
    main_doms: [xml.dom.minidom.Element]  # body直接子节点

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
                recurrent_dom(body_dom)

        print(dom_rely_dict)











if __name__ == '__main__':
    # Args
    root = 'D:\Project\Dataset\排版-完整版财务知识'
    pre_file_name = '...'
    unzip_file_name = ''

    # 1、文件zip->unzip
    # docx_zip_unzip(root, pre_file_name)


    # 2.
    parse_xml_first_tag(root + '\压缩包解压目录\\')






