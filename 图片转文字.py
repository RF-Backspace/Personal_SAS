# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import requests
import base64


def ocr(img_path: str) -> list:
    headers = {
        'Host': 'cloud.baidu.com',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36 Edg/89.0.774.76',
        'Accept': '*/*',
        'Origin': 'https://cloud.baidu.com',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://cloud.baidu.com/product/ocr/general',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    }
    with open(img_path, 'rb') as f:
        img = base64.b64encode(f.read())
    data = {
        'image': 'data:image/jpeg;base64,'+str(img)[2:-1],
        'image_url': '',
        'type': 'https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic',
        'detect_direction': 'false'
    }
    response = requests.post(
        'https://cloud.baidu.com/aidemo', headers=headers, data=data)

    ocr_text = []
    result = response.json()['data']
    if not result.get('words_result'):
        return []

    for r in result['words_result']:
        text = r['words'].strip()
        ocr_text.append(text)
    return ocr_text


'''
img_path 为图片位置，根据需求进行修改
生成的txt文件与代码文件位于同一路径
'''
img_path = 'C:\\Users\\xiaobin.ma\\Desktop\\飞书20220105-102827.jpg'
# content 是识别后得到的结果
content = "".join(ocr(img_path))
# 输出结果
print(content)

with open('结果.txt','a+',encoding='utf-8') as f:
    f.write(content)
    f.write('\r\n')
    f.close()