import openpyxl
import webbrowser
from urllib.parse import quote
import time

# 配置
EXCEL_PATH = "上海、深圳、杭州、北京、苏州.xlsx"  # 替换为你的Excel路径
SEARCH_DELAY = 1  # 搜索间隔时间

def show_sheet_list(wb):
    """显示所有工作表列表"""
    print("\n可用工作表列表：")
    for i, sheet in enumerate(wb.worksheets, 1):
        print(f"[{i}] {sheet.title}")

def get_sheet_range(wb):
    """获取用户选择的工作表范围"""
    total = len(wb.worksheets)
    show_sheet_list(wb)
    
    while True:
        try:
            start = int(input(f"\n请输入起始工作表编号 (1-{total})："))
            count = int(input(f"请输入要处理的工作表数量 (1-{total - start + 1})："))
            
            if 1 <= start <= total and count >= 1:
                end = min(start + count - 1, total)
                return (start-1, end)  # 转换为0-based索引
            print("输入超出有效范围，请重新输入！")
        except ValueError:
            print("请输入有效数字！")

def process_excel():
    """主处理函数"""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    
    # 获取用户选择范围
    start_idx, end_idx = get_sheet_range(wb)
    
    # 处理选定范围的工作表
    for sheet in wb.worksheets[start_idx:end_idx+1]:
        print(f"\n正在处理工作表：{sheet.title}（{sheet.parent.index(sheet)+1}/{len(wb.worksheets)}）")
        
        # 遍历B3-B45单元格
        for row in range(3, 46):
            cell = sheet[f'B{row}']
            if cell.value and str(cell.value).strip():
                keyword = str(cell.value).strip()
                print(f"正在搜索：{keyword}")
                
                # 生成并打开搜索链接
                search_url = f"https://www.baidu.com/s?wd={quote(keyword)}"
                webbrowser.open_new_tab(search_url)
                time.sleep(SEARCH_DELAY)

if __name__ == "__main__":
    process_excel()