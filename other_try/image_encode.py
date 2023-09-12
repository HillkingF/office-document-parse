"""
-*- coding: utf-8 -*-
Time     : 2023/9/7
Author   : Hillking
File     : parse_image.py
Function : 图像分析
"""


def image_to_base64(image_path):
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read())
            return encoded_string.decode("utf-8")  # 将bytes转换为字符串



if __name__ == '__main__':
    import base64

    # 使用示例
    image_path = "C:\\Users\\fengwenni001\Desktop\壁纸\\1.jpg"
    base64_data = image_to_base64(image_path)
    print(base64_data)



