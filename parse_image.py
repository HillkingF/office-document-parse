"""
-*- coding: utf-8 -*-
Time     : 2023/9/7
Author   : Hillking
File     : parse_image.py
Function : 图像分析
"""
import os.path
from xml.dom.minidom import Element
import pickle
import base64

TYPE_IMAGE = 'I'

class ImageParse:
    def __init__(self):
        pass

    def image_serialize(self, image_path):
        if not os.path.exists(image_path):
            raise (2, 'No image path: ' + image_path)
        image_type = image_path.split('.')[-1]

        # 读取图片
        with open(image_path, 'rb') as image_file:
            image_data = image_file.read()
        # 序列化图片
        serialized_image = base64.b64encode(pickle.dumps(image_data)).decode('utf-8')

        # 保存序列化后的图像
        # with open('serialized_image.txt', 'wb') as serialized_image_file:
        #     serialized_image_file.write(serialized_image.encode())
        return image_type, serialized_image

    def get_image_path(self, root_dom: Element, annex_relation):
        root_dom_name = root_dom.nodeName
        if root_dom_name == 'w:drawing':
            # 获取图像编号
            pic_code_doms = root_dom.getElementsByTagName('a:blip')
            if len(pic_code_doms) != 1:
                raise ('图像标签存在多个子 <a:blip>标签！')
            pic_code = pic_code_doms[0].getAttribute('r:embed')
            attr_rela = annex_relation.get(pic_code)
            attr_type = attr_rela.get('Type')
            attr_name = attr_rela.get('Target').split('/')[-1]
            if attr_name.split('.')[-1] in ['png', 'jpg', 'emf']:
                file_type = TYPE_IMAGE
                file_name = attr_name
            else:
                raise('非图片或其他类型图片！')
            return file_type, file_name


