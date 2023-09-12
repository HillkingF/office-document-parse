"""
-*- coding: utf-8 -*-
Time     : 2023/8/23
Author   : Hillking
File     : main.py
Function : docx文档解析

https://docs.python.org/zh-cn/3/library/xml.dom.minidom.html
"""
import os.path
from xml.dom.minidom import parse
from parse_text import TextParse
from merge_mark2txt import MergeParseAndMark
import trans_docx2xml
import shutil
import pickle

class DocxParse:
    def __init__(self, root: str='',  docx_dir_name: str='', mark_dir_path: str=''):
        '''
        :para root: 根目录路径
        :para docx_file: 包含docx文件的目录名称

        :var self.root: 所有文档的根目录，用于保存文档，创建缓存目录(验证路径存在)
        :var self.docx_path: 包含docx文件的目录名称（验证路径存在）
        :var self.mark_dir_path: 标注文档完整路径
        '''

        if not os.path.exists(root):
            raise (2, 'No such dir: ' + root)
        self.root = root

        docx_path = os.path.join(root, docx_dir_name)
        if not os.path.exists(docx_path):
            raise (2, 'No such docx path: ' + docx_path)
        self.docx_path = docx_path

        if not os.path.exists(mark_dir_path):
            raise (2, 'No such docx path: ' + mark_dir_path)
        self.mark_dir_path = mark_dir_path



    def get_n_triplet(self):
        unzip_file_path = os.path.join(self.root, 'unzip')  # 解压目录
        annex_dir = os.path.join(self.root, 'Annexdir')    # 附件保存目录
        triplet_dir = os.path.join(self.root, 'triplet')   # 解析后的元组保存目录
        if not os.path.exists(annex_dir):
            os.mkdir(annex_dir)
        if not os.path.exists(triplet_dir):
            os.mkdir(triplet_dir)

        trans_docx2xml.docx_zip_unzip(self.root, self.docx_path, unzip_file_path)

        docx_files = os.listdir(self.docx_path)
        for docx_file_name in docx_files:
            # print(docx_file_name)
            unzip_dir_name = docx_file_name.replace('.docx', '')


            docu_xml = os.path.join(unzip_file_path, unzip_dir_name, 'word', 'document.xml')
            numb_xml = os.path.join(unzip_file_path, unzip_dir_name, 'word', 'numbering.xml')

            docu_xml_rels = os.path.join(unzip_file_path, unzip_dir_name, 'word')
            final_annex_dir = os.path.join(annex_dir, unzip_dir_name)
            dom_xml = parse(docu_xml)
            root_dom = dom_xml.documentElement  # 获取文档唯一的根节点 <w:document>
            text_parse = TextParse(numb_xml, docu_xml_rels, final_annex_dir)
            text_parse.parse_different_dom(root_dom)  # 根节点初始化
            para_lst = text_parse.para_info_lst


            # mark file parse
            mark_file_path = os.path.join(self.mark_dir_path, unzip_dir_name+'.txt')

            if not os.path.exists(mark_file_path):
                # print('No such mark file: ' + mark_file_path)
                pass
            else:
                # print(mark_file_path)
                merge = MergeParseAndMark(para_lst, mark_file_path)
                merge.build_index_content()
                index_and_content = merge.get_index_triplet()

                # 序列化保存该四元组列表
                triplet_file = os.path.join(triplet_dir, unzip_dir_name + '.pickle')
                with open(triplet_file, 'wb') as f1:
                    pickle.dump(index_and_content, f1)


        # 删除压缩后的docx文件
        shutil.rmtree(unzip_file_path)

        return triplet_dir, annex_dir





if __name__ == '__main__':
    '''
    目录结构：
    - root
        - docx_dir_name
        - annex_dir_name
        - triplet_dir_name
    '''

    root = 'D:\Project\Dataset\排版-完整版财务知识'
    docx_dir_name = '排版'
    annex_dir_name = 'D:\Project\Dataset\排版-完整版财务知识\标注文档'


    tools = DocxParse(root, docx_dir_name, annex_dir_name)
    triplet_dir, annex_dir = tools.get_n_triplet()
    print(triplet_dir)
    print(annex_dir)

    # with open(r'D:\Project\Dataset\排版-完整版财务知识\triplet\2 贝壳集团核算月结管理制度.pickle', 'rb') as f1:
    #     x = pickle.load(f1)
    #     print(x)