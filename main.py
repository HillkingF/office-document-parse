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
    


    # 解析XML文件
    tree = ET.parse(xml_dir)
    root = tree.getroot()

    # 遍历XML文档
    for elem in root.iter():
        print(elem.tag, elem.attrib)

    # 获取指定标签的元素值
    for elem in root.iter('tag_name'):
        print(elem.text)




if __name__ == '__main__':
    # Args
    root = 'D:\Project\Dataset\排版-完整版财务知识\\test'

    # 文件zip->unzip
    docx_zip_unzip(root, '测试文件')




