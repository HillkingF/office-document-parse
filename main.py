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
from xml.dom.minidom import parse   ## https://docs.python.org/zh-cn/3/library/xml.dom.minidom.html

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

# 遍历所有的xml文档，分析同级标签
def parse_xml_first_tag(xml_dir:str=''):
    tag_lst = []
    tag_lst_only = []

    # 验证docx的document.xml中的所有标签是否有不同库但同名的情况
    for root, dirs, files in os.walk(xml_dir):
        for dir in dirs:
            document_file = xml_dir + dir + '\\word\document.xml'
            if os.path.exists(document_file):
                # 解析XML文件
                tree = ET.parse(document_file)
                root = tree.getroot()

                # 遍历XML文档
                for elem in root.iter():
                    if elem.tag not in tag_lst:
                        tag_lst.append(elem.tag)
                        tag_name = elem.tag.split('}')[1]
                        # print(elem.tag)
                        tag_lst_only.append(tag_name)
    
    if len(tag_lst) != len(tag_lst_only):
        print('document.xml文件中存在标签同名不同库！')
        print(len(tag_lst))
        print(len(tag_lst_only))
        return


    # 筛选document.xml文件中<w:body>标签内所有同级标签
    for root, dirs, files in os.walk(xml_dir):
        for i, dir in enumerate(dirs):
            document_file = xml_dir + dir + '\\word\document.xml'
            if os.path.exists(document_file):
                # 解析XML文件
                tree = ET.parse(document_file)
                # 判断所有的跟标签是否都为 document
                root = tree.getroot()   # 根标签统一为document
                if root.tag.split('}')[1] != 'document':
                    print(dir)
                    print(root.tag)
                    exit()
                # 解析根标签的子标签
                for child in root.iter('document'):
                    print(child.tag)


                # 遍历XML文档
                # for elem in root.iter():
                #     if elem.tag not in tag_lst:
                #         tag_lst.append(elem.tag)
                #         tag_name = elem.tag.split('}')[1]
                #         # print(elem.tag)
                #         tag_lst_only.append(tag_name)

                # # 获取指定标签的元素值
                # for elem in root.iter('tag_name'):
                #     print(elem.text)









if __name__ == '__main__':
    # Args
    root = 'D:\Project\Dataset\排版-完整版财务知识'
    pre_file_name = '...'
    unzip_file_name = ''

    # 1、文件zip->unzip
    # docx_zip_unzip(root, pre_file_name)


    # 2.
    parse_xml_first_tag(root + '\压缩包解压目录\\')






