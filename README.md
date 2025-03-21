# 安装
```shell
pip install office-tool-plus -i https://mirrors.aliyun.com/pypi/simple/
```

# 使用
## excel
### to_pdf(excel_path: str, sheet_names: list = None, pdf_dir: str = None) 
> 将 Excel 导出为pdf

**参数：**
> - excel_path: str - excel文件路径
> - sheet_names: list - 需要转换
> - pdf_dir: str - pdf保存的文件夹
> - return：生成的pdf文件路径

**示例：**
```shell
from office_tool_plus import excel
# 将整个Excel导出为pdf，并保存在源目录下
excel.to_pdf('test.xlsx')
# 将整个Excel导出为pdf，并保存在output目录下
excel.to_pdf('test.xlsx', pdf_dir='output')
# 将指定sheet导出为pdf，并保存在源目录下
excel.to_pdf('test.xlsx', ['Sheet1', 'Sheet2'])
# 将指定sheet导出为pdf，并保存在output目录下
excel.to_pdf('test.xlsx', ['Sheet1', 'Sheet2'], 'output')
```


