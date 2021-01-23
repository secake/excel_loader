# 一个基于python3，openpyxl的excel数据加载，导入，存储，导出的引擎  
##  能力  
### 导入能力  
    1. 支持导入多个表格  
    2. 支持指定标题所在行  
    4. 支持指定从哪一行开始解析，如历史解析到了50行，这次可以指定从51行开始解析  
### 映射能力  
    1. 支持excel中一行对应库中多表  
    2. 支持数据某一字段，从数据其他字段进行正则匹配， 例如，书籍的作者指定为作者的id，省市从详细地址中进行正则匹配等等操作  
    3. 支持数据某一字段 从excel多列加载， 例如可以实现数据中有个备注字段，excel中作者，城市以累加的方式 存储到备注字段 
    4. 支持进行数据字段映射，如excel中 90-100分 对应库中优
    5.  支持指定某些值被忽略掉， 如none, -, 无, 否, 非, test等等  
    6.  支持指定字段必填。必填字段为空，则此行数据被抛弃  
### 导出能力  
    1. 可以将加载后数据导出成json格式  
    2. 可以将加载后数据导出成object格式  
    3. 可以将数据存储到库中。
    4. 存储过程中可以指定唯一键，同时可以指定当唯一键发生冲突时的处理方式
    5. 可以将过程中的日志导出  
## 安装openpyxl  
    pip install openpyxl  
## 下载excel_loader  
    git clone https://github.com/SkyingzZ/excel_loader.git  
## 目录结构  
```
excel_loader  
├─── example/student
│       ├───  example.json  excel解析策略 
│       ├───  example.xlsx  excel源文件 
│       └───  example.py    示例代码文件   
│
├─── log.py 日志库代码  
└─── excel_loader.py 核心源代码
```  
### 鼓励贡献代码 
### 使用中有问题请提到issues  
## 使用  
基础使用
```python
# 加载配置以及数据
loader = Loader(config='example.json', path='example.xlsx', globals=globals())

# 以json格返回数据
print(loader.out_json_str())
# 获取加载好的object数据
objs = loader.out_objs()

# 保存数据到库中，需要加载的类继承自 django.db.models.Model
loader.save()
```
#### 进阶使用  
加载多个文件
```python
# 加载配置以及数据
loader = Loader(config='example.json', globals=globals())

# 加载多个文件
loader.load('example.xlsx')
loader.load('example2.xlsx')
loader.load('example3.xlsx')

# 保存数据到库中，需要加载的类继承自 django.db.models.Model
loader.save()
```
指定日志输出
```python
# 指定日志等级
log = Log(Log.Level.INFO)

# 加载配置以及数据
loader = Loader(config='example.json', log=log, globals=globals())
loader.load('example.xlsx')

# 获取字符串形式日志
print(log.read())
```

## 配置  
1. 最基础的配置方式  
```json
{
    "sheets": [
        {"sheet_name": "书籍信息表"}
    ],
    "maps": {
        "Book.name": {"headers": ["书名"]},
        "Book.price": {"headers": ["价格"]},
        "Book.author": {"headers": ["作者"]},
        "Book.publish": {"headers": ["出版时间"]}
    }
}

```
如上配置了 加载一个"书籍信息表"， 其中的书名对应到库中的Book.name，其中Book是库中的类型, name是Book的字段名，价格对应到库中的Book.price  
2. 进阶  
```json
{
    "sheets": [
        {
            "sheet_name": "书籍信息表",
            "header_line": 1,
            "start_line": 1,
            "ignore_values": ["-"]
        },
    ],
    "maps": {
        "Book.name": {
            "headers": ["书名"],
            "required": true,
            "unique": true,
            "conflict": "ignore"
        },
        "Book.price": {"headers": ["价格"]},
        "Book.author": {
            "headers": ["作者"],
            "values": {
                "莫先生": "莫言",
                "$d": "$s"
            }
        },
        "Book.publish": {"headers": ["出版时间"]}
    }
}
```
如上配置了 指定了"书籍信息表"标题所在行，以及从哪行开始解析，和忽略excel中值为-的数据，其中Book.name是必填项，并且唯一，当Book.name和库中数据冲突时，忽略冲突，不保存新数据， author字段进行了值的映射，当excel中作者时莫先生时，转换为莫言。\$d是特殊变量，后面解释  
3. 继续进阶  
```json
{
    "sheets": [
        {"sheet_name": "书籍信息表"}
    ],
    "maps": {
        "Book.name": {"headers": ["书名"]},
        "Book.price": {"headers": ["价格"]},
        "Book.author": {"headers": ["作者"]},
        "Book.remark": {
            "headers": ["出版时间","页数"],
            "func": "strappend"
        },
        "Book.category": {
            "headers": ["分类"],
            "values": [
                "$r:(^[\\w]+$)$g:(0)":"$s",
                "$r:([\\w]+)$g:(0)":"$r:([\\w]+)$g:(0)",
                "$d": "其他"
            ]
        },
    }
}
```
如上配置了 将表格中的出版时间和页数以字符串累加的方式赋值到Book.remark字段，同时分类进行了值的映射  
### 映射关系说明  
"values": [
    //如果匹配，赋值为右面  
    "xx": "xx"
]
### 映射特殊变量说明  
```
$f:(class.field)    指定从那个类的那字段映射数据，默认是当前类当前字段  
$r:(pattern)        指定映射的方式，取正则匹配的部分,正则默认 .*  
$g:(0)              必须出现$r:(pattern)时可用。为取正则结果的哪一个分组，默认0  
$d                  指定default的值。即再所有values映射中都没匹配的。！（当配置没有values时，$d默认是$s，当有values并且没有出现$d时，$d是None）  
#s                  代表 self
```
### 特殊函数说明  
```
listappend  以累加方式加入list
strappend   以累加方式加入字符串
setadd      以add方式加入set
numadd      以add方式加 int 或float
```

