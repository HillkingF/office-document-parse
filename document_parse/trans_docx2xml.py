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

def docx_zip_unzip(root: str='', docx_files_dir:str='', unzip_file_path: str=''):
    # 遍历所有的子文档，进行压缩，暂存后解压
    predir = docx_files_dir
    zipdir = os.path.join(root, 'zip_cache')    # docx 2 zip dir


    if not os.path.exists(zipdir):
        os.mkdir(zipdir)
    if not os.path.exists(unzip_file_path):
        os.mkdir(unzip_file_path)

    files = os.listdir(predir)
    for file in files:

        file_name = file.replace('.docx', '')
        old_file_path = os.path.join(predir, file)
        new_file_path = os.path.join(zipdir, file_name + '.zip')
        shutil.copy(old_file_path, new_file_path)  # 将所有的docx文档转换成zip格式的文件

    zip_files = os.listdir(zipdir)
    for zip_file in zip_files:
        file_name = zip_file.replace('.zip','')
        zip_file_path = os.path.join(zipdir, zip_file)
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # 解压缩文件到指定目录
            zip_ref.extractall(os.path.join(unzip_file_path,file_name))

    # 删除缓存文件
    shutil.rmtree(zipdir)

