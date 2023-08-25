"""
-*- coding: utf-8 -*-
Time     : 2023/8/23
Author   : Hillking
File     : main.py
Function : docx文档解析

https://docs.python.org/zh-cn/3/library/xml.dom.minidom.html
"""

import json
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.dom.minidom import parse   ##
import xml.dom.minidom as xdm
import re
from numering_parse import NumberParse

number_parse = NumberParse()
number_parse.get_numbering('D:\Project\Dataset\排版-完整版财务知识\压缩包解压目录\\2 贝壳集团核算月结管理制度\word\\numbering.xml')
# print(number_parse_exp.abstractnum_dict)
# print(number_parse_exp.numid_absid_dict)
# print(number_parse_exp.num_pointer)
# exit()




# 全局变量库
dom_rely_dict: {str:list} = {}


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


def set_recurrent_dom(root_dom: xdm.Element) -> None:
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

def get_recurrent_dom(root_dom: xdm.Element) -> None:
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
    tag_lst_only = []
    main_doms: [xdm.Element]  # body直接子节点

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

        # print(dom_rely_dict)
        # dom_json = json.dumps(dom_rely_dict)
        # with open("dom.json", 'a', encoding='utf-8') as jsonw:
        #     jsonw.write(dom_json)




def get_content_from_tag_p(root_dom:xdm.Element, p_text:str='') -> str:
    # TODO: 修改段落文本解析代码
    '''
    :param root_dom: 节点：不为空节点
    :param p_text: 各阶段的文本
    :return: 返回最终段落的文本
    :func: 解析p标签内部子节点
    '''

    if root_dom is None:
        return p_text

    root_dom_name  = root_dom.nodeName
    if root_dom_name == 'w:t':
        dom_text = root_dom.childNodes[0].data
        p_text = p_text + dom_text

    if root_dom_name == 'w:numPr':  # 编号解析
        num_text = number_parse.get_parse_number_text(root_dom)
        p_text = p_text + num_text + ' '
        if p_text == '非中文类型文本暂未解析':
            pass


    child_doms = root_dom.childNodes
    if len(child_doms) ==0:
        return p_text

    for child_dom in child_doms:
        p_text = get_content_from_tag_p(child_dom, p_text)
    return p_text


def get_content_from_tag_tb(root_dom:xdm.Element) -> [str, [str]]:
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
        if child_dom_name == 'w:tr':   # 解析每一行的元素
            tb_cel_doms = child_dom.childNodes
            row_text = ''  # 每一行的文本初始化
            for tb_cel_dom in tb_cel_doms:
                tb_cel_dom_name = tb_cel_dom.nodeName
                if tb_cel_dom_name == 'w:tc': # 解析每一行的每一列元素
                    tb_p_doms = tb_cel_dom.getElementsByTagName('w:p')
                    p_text = ''
                    for tb_p_dom in tb_p_doms:
                        p_text = get_content_from_tag_p(tb_p_dom, p_text)
                    p_text = ' ' if p_text == '' else p_text
                    row_text = row_text + ' | ' + p_text
            row_text = row_text + ' | \n'
            tb_whole_text = tb_whole_text + row_text
            tb_row_lst.append(row_text)

    return tb_whole_text, tb_row_lst



def get_content_from_tag_sdt(root_dom:xdm.Element):
    '''
    :func: 从结构体中解析特定文本
    :param root_dom:
    :return:
    '''

    # TODO：补充代码
    pass

def get_text_from_dom(xml_path:str=''):
    dom_xml = parse(xml_path)
    root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>
    root_dom_name = root_dom.nodeName

    get_content_from_tag_p(root_dom)





def parse_different_dom(root_dom:xdm.Element):
    '''
    :param root_dom: 节点
    :return:
    :func: 不同类型的节点判断，并选择不同的节点处理方法

    节点分级：
    body子节点：p\tbl\sdt
    p内子节点：
    '''

    if root_dom is None:
        return



    if root_dom.nodeName == 'w:p':
        p_text = get_content_from_tag_p(root_dom, '')  # 文本段落处理
        print(p_text)
    elif root_dom.nodeName == 'w:tbl':
        get_content_from_tag_tb(root_dom)
    elif root_dom.nodeName == 'w:sdt':
        get_content_from_tag_sdt(root_dom)
    # 待扩展的标签：图片、附件
    # elif root_dom.nodeName == '':


    # 子节点循环遍历
    child_doms = root_dom.childNodes
    if len(child_doms) == 0:
        return
    for child_dom in child_doms:
        parse_different_dom(child_dom)









if __name__ == '__main__':
    # Args
    root = 'D:\Project\Dataset\排版-完整版财务知识'
    pre_file_name = '...'
    unzip_file_name = ''

    # 1、文件zip->unzip
    # docx_zip_unzip(root, pre_file_name)


    # 2.遍历所有的xml文档，將所有的节点依赖关系保存到字典及文件中
    # parse_xml_first_tag(root + '\压缩包解压目录\\')


    # 3.查看关键标签
    # get_and_analyse_dom('w:p')

    # 4.获取标签中的全部文本
    # get_text_from_dom('D:\Project\Dataset\排版-完整版财务知识\压缩包解压目录\\2 贝壳集团核算月结管理制度\word\document.xml')

    # 6.解析所有的标签，获取文本内容
    dom_xml = parse('D:\Project\Dataset\排版-完整版财务知识\压缩包解压目录\\2 贝壳集团核算月结管理制度\word\document.xml')
    root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>
    parse_different_dom(root_dom)

