"""
-*- coding: utf-8 -*-
Time     : 2023/8/28
Author   : Hillking
File     : parse_annex.py
Function : parse annexs, include:  image、excel
"""


import os.path
from xml.dom.minidom import parse
import shutil

class AnnexParse:

    def __init__(self, annex_dir):
        annex_rels_path = os.path.join(annex_dir, '_rels', 'document.xml.rels')
        if not os.path.exists(annex_rels_path):
            raise os.error(2, 'No such file named \'' + annex_rels_path + '\'')
        dom_xml = parse(annex_rels_path)
        self.root_dom = dom_xml.documentElement  # Get Only one root dom：<Relationships>
        if self.root_dom is None or self.root_dom.nodeName != 'Relationships':
            raise os.error(2, '\'document.xml.rels\' file root dom error!')

        ## 外部可访问属性：
        self.annex_relation: {str: {str: str}} = {}
        self.annex_media_dir = os.path.join(annex_dir, 'media')
        self.annex_embeddings_dir = os.path.join(annex_dir, 'embeddings')

    def get_annex_relation(self):
        child_doms = self.root_dom.childNodes
        if len(child_doms) == 0:
            return self.annex_relation  # No annex relationship

        for i, child_dom in enumerate(child_doms):
            child_dom_name = child_dom.nodeName
            if child_dom_name != 'Relationship' :
                raise os.error(2,'Annex parse error! Different child dom like <Relationship> in tag Relationships')  # Un-chick relationship, need to deal!

            Id = child_dom.getAttribute('Id')
            Target = child_dom.getAttribute('Target')
            Type = child_dom.getAttribute('Type')

            attr: {str: str}={}
            attr['Target'] = Target.split('/')[-1]
            attr['Type'] = Type
            self.annex_relation[Id] = attr

        return self.annex_relation

    def copy_annex(self, files_path):
        '''
        :func : 将附件移动到新的目录中便于访问
        :return:
        '''
        for _, _, files in os.walk(self.annex_media_dir):
            for file in files:
                shutil.copy(os.path.join(self.annex_media_dir, file), os.path.join(files_path, file))

        for _, _, files in os.walk(self.annex_embeddings_dir):
            for file in files:
                shutil.copy(os.path.join(self.annex_embeddings_dir, file), os.path.join(files_path, file))





if __name__ == '__main__':
    annex_dir = 'xxx\word'
    annex_example = AnnexParse(annex_dir)
    annex_relation = annex_example.get_annex_relation()

    annex_final_dir = final_annex_dir = 'xxx\\10 花桥学院核算指引'
    annex_example.copy_annex(annex_final_dir)







