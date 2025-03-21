from pathlib import Path


class WinExcel:
    def __init__(self):
        from win32com.client import gencache
        self.app = gencache.EnsureDispatch("Excel.Application")
        self.app.Visible = False

    @staticmethod
    def check_path(input_path: str, suffix: str, output_dir: str = None):
        """
        检查给定的文件路径是否存在。如果提供了文件保存的目录，则检查该目录是否存在，
        并根据源文件名在该目录下生成对应的新文件路径。

        :param input_path: 源文件的路径
        :param suffix：文件后缀
        :param output_dir: 输出文件的目录，如果不传，则默认在源文件目录
        :raises FileNotFoundError: 如果指定的文件路径或目录不存在
        """
        # 获取 excel 文件的绝对路径
        try:
            abs_input_path = Path(input_path).resolve(strict=True)
        except FileNotFoundError:
            raise FileNotFoundError(f"指定的文件路径 '{input_path}' 不存在。")

        # 获取 pdf 文件的绝对路径
        if output_dir is None:
            abs_output_dir = abs_input_path.with_suffix('.pdf')
        else:
            abs_pdf_dir = Path(output_dir).resolve()
            if not abs_pdf_dir.exists():
                raise FileNotFoundError(f"指定的保存目录 '{abs_pdf_dir}' 不存在。")
            abs_output_dir = abs_pdf_dir / f"{abs_input_path.stem}.{suffix}"
        return abs_input_path, abs_output_dir

    def to_pdf(self, excel_path: str, sheet_names: list = None, pdf_dir: str = None):
        excel_path, pdf_path = self.check_path(excel_path, "pdf", pdf_dir)
        # 如果没有提供特定的工作表名称，则导出所有工作表
        workbook = self.app.Workbooks.Open(excel_path)
        sheet_names = sheet_names or [sheet.Name for sheet in workbook.Sheets]
        sheets_to_hide = [sheet for sheet in workbook.Sheets if sheet.Name not in sheet_names]
        try:
            # 临时隐藏不需要导出的工作表
            for sheet in sheets_to_hide:
                sheet.Visible = False
            # 导出可见的工作表为PDF
            workbook.ExportAsFixedFormat(0, str(pdf_path))
        finally:
            # 恢复所有工作表的可见性
            for sheet in sheets_to_hide:
                sheet.Visible = True
            workbook.Close(SaveChanges=False)
        self.app.Quit()
        return pdf_path


if __name__ == '__main__':
    excel = WinExcel()
    excel_p = f"E:/devProject/office_tool_plus/tests/印尼出口总箱单及详细箱单模板（元亨冷却塔3.19）(1).xlsx"
    sheet_n = ["156", "157"]
    excel.to_pdf(excel_p, sheet_n)
