
class Student:
    def __init__(self):
        self.age = 0
        self.name = ''
        self.teacher = ''
        self.grade = ''
        self.class_ = ''
        self.score = 0
    def save(self):
        print('save')

class Score:
    def __init__(self):
        self.score = 0
        self.commit = ''
    def save(self):
        print('save')

import os
import sys

base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(base_path)
from excel_loader import *


if __name__ == "__main__":
    # 用法 1
    # 加载配置
    # loader = Loader('example.json', globals=globals())
    # 加载数据
    # loader.load('example.xlsx')
    # 以json格返回数据
    # print(loader.out_json_str())

    #用法2
    # # 加载配置以及数据
    # loader = Loader('example.json', path='example.xlsx', globals=globals())
    # # 以json格返回数据
    # print(loader.out_json_str())
    
    # 用法3


    from log import Log
    # 指定日志等级
    log = Log(Log.Level.INFO)
    loader = Loader(base_path+'/example/student/example.json', path=base_path+'/example/student/example.xlsx', log=log,  globals=globals())
    # # 可以继续加载其他的 excel
    # loader.load('example.xlsx')
    print(loader.out_json_str())
    # # 可以获取日志
    # print(log.read())
    
    # 可以返回加载后的obj格式数据
    # print(loader.out_objs())

    # 保存数据到库中， 需要继承django.db.models.Model
    # loader.save()