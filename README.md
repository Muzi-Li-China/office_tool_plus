# 安装

> `office-tool-plus` 一些自动化办公的工具包

```shell
# 国内源
pip install --upgrade office-tool-plus -i https://mirrors.aliyun.com/pypi/simple/ 
# 官方源（最新版）
pip install --upgrade office-tool-plus -i https://pypi.org/simple  
```

# 使用

## excel

### single_to_pdf(excel_path, sheet_names=None, pdf_dir=None):

> 将指定`Excel`工作簿中的工作表导出为`PDF`格式。

**参数：**

- `excel_path`: `Excel`文件的路径。
- `sheet_names`: 需要导出的工作表名称列表。如果未提供，则默认导出所有工作表。
- `pdf_dir`: `PDF`文件的保存目录。如果未提供，则默认保存在`Excel`文件的同目录下。

**返回：**

- `pdf_path`: 导出的`PDF`文件路径。

**示例：**

```python
from office_tool_plus import excel

# 将整个 Excel 导出为 pdf，并保存在源目录下
excel.single_to_pdf('test.xlsx')
# 将整个 Excel 导出为 pdf，并保存在 output 目录下
excel.single_to_pdf('test.xlsx', pdf_dir='output')
# 将指定 sheet 导出为 pdf，并保存在源目录下
excel.single_to_pdf('test.xlsx', ['Sheet1', 'Sheet2'])
# 将指定 sheet 导出为 pdf，并保存在 output 目录下
excel.single_to_pdf('test.xlsx', ['Sheet1', 'Sheet2'], 'output')
```

### many_to_pdf(excel_dir, suffix=None, recursive=True, pdf_dir=None):

> 将指定目录下的`Excel`文件批量导出为`PDF`格式。

**参数：**

- `excel_dir`: `Excel`文件的目录路径。
- `suffix`: 需要导出的文件后缀名列表。如果未提供，则默认导出`["*.xlsx", "*.xls"]`
- `recursive`: 是否递归搜索子目录。如果为`True`，则递归搜索子目录，否则只搜索当前目录。
- `pdf_dir`: `PDF`文件的保存目录。如果未提供，则默认保存在`Excel`文件的同目录下。

**返回：**

- `None`

**示例：**

```python
from office_tool_plus import excel

# 将 test 目录下（包含子目录）所有的 Excel 文件批量导出为 pdf，并保存在源目录下
excel.many_to_pdf('test')
# 将 test 目录下（包含子目录）所有的 Excel 文件批量导出为 pdf，并保存在 output 目录下
excel.many_to_pdf('test', pdf_dir='output')
# 将 test 目录下（不包含子目录）所有的 Excel 文件批量导出为 pdf，并保存在源目录下
excel.many_to_pdf('test', recursive=False)
# 将 test 目录下（不包含子目录）,后缀是 *.xlsx 的 Excel 文件批量导出为 pdf，并保存在 output 目录下
excel.many_to_pdf('test', suffix=['*.xlsx'], recursive=False, pdf_dir='output')
```

### from_template(template_file, labor_datas, sheet_name=None, output_dir=None):

> 根据模板文件和模板数据生成新的 Excel 文件。

**参数：**

- `template_file`: 模板文件地址
- `labor_datas`: 模板文件中的变量数据，以`[{},{}]`形式传入
    1. `labor_datas` 是一个列表，列表中的每个元素都是一个字典
    2. 例如：
  ```python
  labor_datas = [
      {
          "文件名": "铝材合格证",
          "数据": [
              {"A2": {"type": "text", "value": "铝材合格证"}},
              {"B3": {"type": "text", "value": "AD50-03"}},
              {"B4": {"type": "text", "value": "50平开扇"}},
              {"B5": {"type": "text", "value": "白亮光/TMP106/H0YA-0017A"}},
              {"B6": {"type": "number", "value": 6}},
              {"E3": {"type": "image", "value": "./照片/AD50-03.png"}},
          ]
      }
  ]
  ```
- `sheet_name`: 指定模板工作表，默认为第一个工作表。
- `output_dir`: 输出文件保存的目录。如果未提供，则默认保存在模板文件的同目录下。

**返回：**

- `output_file_list`：生成的 Excel 文件路径列表。

**示例：**

```python
from office_tool_plus import excel

labor_datas = [
    {
        "文件名": "铝材合格证",
        "数据": [
            {"A2": {"type": "text", "value": "铝材合格证"}},
            {"B3": {"type": "text", "value": "AD50-03"}},
            {"B4": {"type": "text", "value": "50平开扇"}},
            {"B5": {"type": "text", "value": "白亮光/TMP106/H0YA-0017A"}},
            {"B6": {"type": "number", "value": 6}},
            {"D6": {"type": "text", "value": "6063-T5"}},
            {"B7": {"type": "number", "value": 150}},
            {"B7": {"type": "number", "value": 150}},
            {"D7": {"type": "number", "value": 0.07}},
            {"E3": {"type": "image", "value": "./照片/AD50-03.png"}},
            {"F9": {"type": "date", "value": time.strftime("%Y-%m-%d", time.localtime())}},
        ]
    },
    {
        "文件名": "型材合格证"
    },
    {
        "数据": [
            {"A2": {"type": "text", "value": "板材合格证"}},
            {"B3": {"type": "text", "value": "PYM-5501AH"}},
            {"B4": {"type": "text", "value": "30框小面"}},
            {"B5": {"type": "text", "value": "白亮光/TMP106/H0YA-0017A"}},
            {"B6": {"type": "number", "value": 6}},
            {"D6": {"type": "text", "value": "6063-T5"}},
            {"B7": {"type": "number", "value": 150}},
            {"B7": {"type": "number", "value": 150}},
            {"D7": {"type": "number", "value": 0.07}},
            {"E3": {"type": "image", "value": r"E:\PythonWork\照片\PYM-5501AH.jpeg"}},
            {"F9": {"type": "date", "value": time.strftime("%Y-%m-%d", time.localtime())}},
        ]
    },
]
template_file = '产品合格证模板.xlsx'
# 根据模板，生成指定的 Word，并保存在 output 目录下
output_file_list = excel.from_template(template_file, labor_datas, output_dir="./output")
```

**效果：**

![img.png](static/img_1.png)

## word

### single_to_pdf(word_path, pdf_dir=None):

> 将指定的`Word`文档导出为`PDF`格式。

**参数：**

- `word_path`: `Word`文件的路径。
- `pdf_dir`: `PDF`文件的保存目录。如果未提供，则默认保存在`Word`文件的同目录下。

**返回：**

- `pdf_path`: 导出的`PDF`文件路径。

**示例：**

```python
from office_tool_plus import word

# 将指定的 Word 导出为 pdf，并保存在源目录下
word.single_to_pdf('test.docx')
# 将指定的 Word 导出为 pdf，并保存在 output 目录下
word.single_to_pdf('test.docx', pdf_dir='output')
```

### many_to_pdf(word_dir, suffix=None, recursive=True, pdf_dir=None):

> 将指定目录下的`Word`文件批量导出为`PDF`格式。

**参数：**

- `word_dir`: `Word`文件的目录路径。
- `suffix`: 需要导出的文件后缀名列表。如果未提供，则默认导出`["*.docx", "*.doc"]`
- `recursive`: 是否递归搜索子目录。如果为`True`，则递归搜索子目录，否则只搜索当前目录。
- `pdf_dir`: `PDF`文件的保存目录。如果未提供，则默认保存在`Word`文件的同目录下。

**返回：**

- `None`

**示例：**

```python
from office_tool_plus import word

# 将 test 目录下（包含子目录）所有的 Word 文件批量导出为 pdf ，并保存在源目录下
word.many_to_pdf('test')
# 将 test 目录下（包含子目录）所有的 Word 文件批量导出为 pdf，并保存在 output 目录下
word.many_to_pdf('test', pdf_dir='output')
# 将 test 目录下（不包含子目录）所有的 Word 文件批量导出为 pdf，并保存在源目录下
word.many_to_pdf('test', recursive=False)
# 将 test 目录下（不包含子目录）,后缀是 *.docx 的 Word 文件批量导出为 pdf，并保存在 output 目录下
word.many_to_pdf('test', suffix=['*.docx'], recursive=False, pdf_dir='output')
```

### from_template(template_file, labor_datas, output_dir=None):

> 根据指定的模板文件生成 Word

**参数：**

- `template_file`: 模板文件地址
- `labor_datas`: 模板文件中的变量数据，以`[{},{}]`形式传入
    1. `labor_datas` 是一个列表，列表中的每个元素都是一个字典，字典的键是模板文件中的变量名，值是变量的值。
    2. 字典中的键名必须与模板文件中的变量名保持一致。
    3. 最终生成的 word 文件名称，优先取字典中的"文件名"字段，如果不存在，则取字典中的"姓名"字段。，如果仍然不存在，则用模板文件名
    4. 如果模板中有照片，则需要传入照片地址，支持相对路径和绝对路径
    5. 所有照片的模板名必须以"照片"结尾
- `output_dir`: 输出文件保存的目录。如果未提供，则默认保存在模板文件的同目录下。

**返回：**

- `output_file_list`：生成的 Word 文件路径列表。

**示例：**

```python
from office_tool_plus import word

template_file = f'准考证模板.docx'
labor_datas = [
    {'姓名': "王乐",
     '性别': "女",
     "身份证号": "421023200001018592",
     "准考证号": "000001",
     "报考岗位": "全栈工程师",
     "头像_照片": "./照片/王乐.png",
     "内容_照片": r"E:\PythonWork\照片\王乐.png",
     "内容": """
     考生须知：
    1.考试当天，考生凭准考证、有效居民身份证（原件），方可正常参加考试。身份证于2024年4月27日（含）之前过期的视为无效身份证，持身份证复印件、户口薄、户籍证明、电子或纸质身份证明不准参加考试。
    2.考生自备橡皮、2B铅笔、黑色签字笔等文具。开考后考生不得传递任何物品。
    3.严禁携带各种通讯设备（如手机、无线耳机、智能手表、运动手环等）、具有计算存储功能的电子设备及与考试相关的资料进入考场。
    4.考试开始前30分钟允许考生进入考场，对号入座，并将准考证、身份证、普通手表摘下放在桌面靠过道拐角处；考试开始30分钟后不得进入考点；考试不得提前交卷、退场。
"""
     },
]

# 根据模板，生成指定的 Word，并保存在 output 目录下
output_file_list = word.from_template(template_file, labor_datas, "./output")
```

**效果：**

![img.png](static/img.png)

## Linux

### single_to_pdf(input_path, convert_to="pdf", output_dir=None, java_home=None, lang=None):

> 使用LibreOffice在Linux平台上转换文档格式。 需要安装 apk add libreoffice openjdk8 font-noto-cjk
> - libreoffice ：用于处理Office文件。
> - openjdk8 ：用于运行LibreOffice。
> - font-noto-cjk ：用于支持中文字体。

**参数：**

- `input_path`: 需要转换的文件路径。
- `convert_to`: 转换后的文件格式。
- `output_dir`: 转换后的文件保存的目录。
- `java_home`: （可选）Java安装目录的路径，默认使用`/usr/bin/java`。
- `lang`: （可选）设置LANG环境变量，默认为`zh_CN.UTF-8`。

**返回：**

- `None`

**示例：**

```python
from office_tool_plus import linux

# 将 test.docx 文件导出为 pdf ，并保存在源目录下
linux.single_to_pdf('test.docx')
# 将 test.xlsx 工作簿导出为 pdf ，并保存在 output 目录下
linux.single_to_pdf('test.xlsx', "output")
```

### many_to_pdf(input_dir, suffix, convert_to="pdf", output_dir=None, java_home=None, lang=None, recursive=False):

> 使用LibreOffice在Linux平台上批量转换文档格式。 需要安装 apk add libreoffice openjdk8 font-noto-cjk
> - libreoffice ：用于处理Office文件。
> - openjdk8 ：用于运行LibreOffice。
> - font-noto-cjk ：用于支持中文字体。

**参数：**

- `input_dir`: 需要转换的目录。
- `suffix`: 需要转换的文件后缀名列表。
- `convert_to`: 转换后的文件格式。
- `output_dir`: 转换后的文件保存的目录。
- `java_home`: （可选）Java安装目录的路径，默认使用`/usr/bin/java`。
- `lang`: （可选）设置LANG环境变量，默认为`zh_CN.UTF-8`。

**返回：**

- `None`

**示例：**

```python
from office_tool_plus import linux

# 将 test 目录下（包含子目录）所有 .xlsx，.doc 后缀的文件批量导出为 pdf ，并保存在源目录下
linux.many_to_pdf("test", ['*.xlsx', '*.doc'])
# 将 test 目录下（包含子目录）所有 .xlsx，.doc 后缀的文件批量导出为 pdf ，并保存在 output 目录下
linux.many_to_pdf("test", ['*.xlsx', '*.doc'], "output")
```
