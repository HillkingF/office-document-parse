"""
-*- coding: utf-8 -*-
Time     : 2023/8/28
Author   : Hillking
File     : trans_docx2xml.py
Function : Trans [doxc] to [xml]
"""

import os
import shutil
import zipfile

def docx_zip_unzip(root: str='', pre_files_dir:str='', unzip_file_path: str=''):
    # 遍历所有的子文档，进行压缩，暂存后解压
    predir = root + '\\' + pre_files_dir + '\\'
    zipdir = root + '\压缩包\\'
    unzip_file_path = root + '\压缩包解压目录\\'

    if not os.path.exists(zipdir):
        os.mkdir(zipdir)
    if not os.path.exists(unzip_file_path):
        os.mkdir(unzip_file_path)

    for _, _, files in os.walk(predir):
        for file in files:
            file_name = file.split('.')[0]
            old_file_path = predir + file
            new_file_path = zipdir + file.split('.')[0] + '.zip'
            shutil.copy(old_file_path, new_file_path)  # 将所有的docx文档转换成zip格式的文件
            with zipfile.ZipFile(new_file_path, 'r') as zip_ref:
                # 解压缩文件到指定目录
                zip_ref.extractall(unzip_file_path + file_name)

    # 删除缓存文件
    shutil.rmtree(zipdir)

