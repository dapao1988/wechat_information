#coding=utf8
"""
# Author: Wenbing.Wang
# Email: wangwenbingood1988@gmail.com
# Created Time : Sun Apr  12 20:59:04 2024

# File Name: example.py
# Description:

"""

HEAD_INFO = 'member_info'
HEAD_NAME = 'name'
HEAD_NICK_NAME = 'nick_name'
HEAD_PROVINCE = 'province'
HEAD_CITY = 'city'
HEAD_SEX = 'sex'
HEAD_SIGNATURE = 'signature'
# 当family捞出人数大于等于3时，标记特殊字体和颜色
WARNING_FAMILY_NUBMER = 3

# 定义一个字典，映射汉字数字和阿拉伯数字之间的对应关系
chinese_to_arabic_dict = {
    "一": "1",
    "二": "2",
    "三": "3",
    "四": "4",
    "五": "5",
    "六": "6",
    "七": "7",
    "八": "8",
    "九": "9"
}

def get_next_letter(letter: str, N: int) -> str:
    """
    获取字母 letter 后面第 N 个字母。

    参数:
    letter (str): 起始字母。
    N (int): 要获取的字母的相对位置。

    返回:
    str: 获取到的字母。
    """
    # 获取起始字母的 ASCII 码值
    letter_ascii = ord(letter)
    # 计算目标字母的 ASCII 码值
    target_ascii = letter_ascii + N
    # 将 ASCII 码值转换为字母
    next_letter = chr(target_ascii)
    return next_letter

