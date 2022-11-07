使用```python-docx``` 库，自动化导出符合 成都信息工程大学 龙泉校区 学生社团管理部 要求的活动证明。

# 写在前面

- 嘿嘿，第一次在学习一个库时用 dir 函数
- 第一次尝试根据官方文档，学习这个库
- 第一次通过源代码了解
> 论学习**面向对象**的重要性
  > 
# 说一说项目里面的这些文件和文件夹

## 文件夹：学习与尝试

### 整体说明

最完美的想法是根据模板要求，全部（比如说模板里面的标题、文字、表格）由程序自动生成。但是应该是精度问题（```python-docx```库的长度单位有很多，比如说```Emu```、```Cm```、```Pt```等等等。）所以生成的表格的**列宽**无法得到百分之百一模一样。

所以这文件夹只是最开始的尝试与学习，也确实是基本上掌握了这个库，代码的注释也很详细，但是并没有完美解决问题。

### 文件说明

见名知意

## 文件夹：CUIT_LongQuan_activity _certificate_automation

### 整体说明

主要的程序所在的文件夹。

采用退而求其次的方法，直接填写模板的表格。

大部分代码直接```copy```学习与尝试里面的

### 文件说明

- 文件：CUIT_LongQuan_activity_certificate_automation_main.py

  主文件

- 文件：CUIT_LongQuan_activity_certificate_automation_tools.py

  不是主文件

- 文件：附件一：成都信息工程大学社团活动证明模板(4).docx

  模板文件，在这个基础上进行填写

# 这个库的一些资料

## 官方文档

- https://www.osgeo.cn/python-docx/api/enum/WdLineSpacing.html
- https://python-docx.readthedocs.io/en/latest/

## 其他

- 字号对照表：https://blog.csdn.net/weixin_42651205/article/details/84395904
- 表格边框：https://www.jianshu.com/p/9ad7db7825ba
- 长度单位：https://blog.csdn.net/qq_39147299/article/details/125601616
- 替换：https://zhuanlan.zhihu.com/p/423704887
- 合并到一个 word 文件：https://zhuanlan.zhihu.com/p/400209096

# 备忘录
## 如果定义了文档网格 **好像没有用**
```python
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml 
#取消设置 ”如果定义了文档网格，则对齐到网格”
para._p.get_or_add_pPr().insert(0,parse_xml('<w:snapToGrid {}  w:val="0"/>'.format(nsdecls('w')))) 

#设置 “如果定义了文档网格，则对齐到网格”
para._p.get_or_add_pPr().insert(0,parse_xml('<w:snapToGrid {}  w:val="1"/>'.format(nsdecls('w'))))

#恢复到默认
for i in para._p.pPr:
    if "snapToGrid" in str(i):
        para._p.pPr.remove(i)
```
## 边框
- https://www.jianshu.com/p/9ad7db7825ba
```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        left={"sz": 24, "val": "dashed", "shadow": "true"},
        right={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('left', 'top', 'right', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))
```

## 合并到一个 word 文件

> https://zhuanlan.zhihu.com/p/400209096
>
> 略有 bug，已改

```py
import os

from docx import Document
from docxcompose.composer import Composer

road_docx = r'C:\output_fafa\cs\cs_docx'
road_all = r'C:\output_fafa\cs\cs_docx'
original_docx_path = road_docx
new_docx_path = f'{road_all}/activity_name.docx'


all_file_path = []
for file_name in os.listdir(original_docx_path):
    all_file_path.append(f'{original_docx_path}/{file_name}')

first_document = Document(all_file_path[0])
first_document.add_page_break()
middle_new_docx = Composer(first_document)

for index, word in enumerate(all_file_path[1:]):
    word_document = Document(word)
    if index != len(all_file_path) - 2:
        word_document.add_page_break()
    middle_new_docx.append(word_document)

middle_new_docx.save(new_docx_path)
```

