import jieba
from wordcloud import WordCloud
import numpy as np
from PIL import Image
from matplotlib import colors
import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import os
import re
from collections import Counter

# 设置中文字体支持
matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
matplotlib.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 定义上海市16个区的分组
district_groups = [
    ["黄浦区", "徐汇区", "长宁区", "静安区"],
    ["普陀区", "虹口区", "杨浦区", "浦东新区"],
    ["闵行区", "宝山区", "嘉定区", "金山区"],
    ["松江区", "青浦区", "奉贤区", "崇明区"]
]

# 确保输出文件夹存在
output_dir = '词频与词云图'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

def findifhave(demo, stop):
    for ret in stop:
        if (demo == ret):
            return 'T'

# 获取停用词列表
stop = ['\n']
with open("stop.txt", 'r', encoding='utf-8') as f1:
    for line in f1:
        stop.append(line.replace("\n", ""))

# 设置文件夹路径
docs_folder = "16区各自全文档"
images_folder = "地区抠图"

# 提取区名的函数 - 只保留"XX区"或"XX新区"部分
def extract_district_name(filename):
    # 尝试匹配"XX区"或"XX新区"格式
    match = re.search(r'([^\s]+[区]|[^\s]+新区)', filename)
    if match:
        return match.group(0)
    return os.path.splitext(filename)[0].strip()  # 如果没有匹配到，返回无扩展名的文件名

# 处理所有区的文本并获取词频
def process_all_districts():
    # 获取所有文本文件
    text_files = []
    if os.path.exists(docs_folder):
        text_files = [f for f in os.listdir(docs_folder) 
                     if os.path.isfile(os.path.join(docs_folder, f)) 
                     and (f.endswith('.txt') or f.endswith('.xlsx'))]
    else:
        print(f"警告: 找不到文档文件夹 '{docs_folder}'")
        return {}

    # 创建文件名到图片路径的映射，使用多种可能的匹配方式
    png_mapping = {}
    if os.path.exists(images_folder):
        for png in os.listdir(images_folder):
            if png.lower().endswith('.png'):
                # 存储原始名称映射
                base_name = os.path.splitext(png)[0].strip()
                png_mapping[base_name.lower()] = os.path.join(images_folder, png)
                
                # 存储提取的区名映射
                district_name = extract_district_name(png)
                png_mapping[district_name.lower()] = os.path.join(images_folder, png)

    # 存储所有区的词频数据
    all_districts_data = {}
    
    for text_file in text_files:
        # 跳过Excel文件
        if text_file.lower().endswith('.xlsx'):
            print(f"  跳过Excel文件: {text_file} - 请先转换为文本格式")
            continue
            
        # 获取文件名（不含扩展名）和区名
        file_base = os.path.splitext(text_file)[0].strip()
        district_name = extract_district_name(text_file)
        
        # 尝试多种方式查找匹配的图片
        png_file = None
        
        # 1. 直接匹配
        possible_png = os.path.join(images_folder, file_base + ".png")
        if os.path.exists(possible_png):
            png_file = possible_png
        
        # 2. 使用区名匹配
        elif district_name.lower() in png_mapping:
            png_file = png_mapping[district_name.lower()]
        
        # 3. 尝试不区分大小写、去除空格的匹配
        elif file_base.lower().replace(" ", "") in [k.lower().replace(" ", "") for k in png_mapping.keys()]:
            for k, v in png_mapping.items():
                if k.lower().replace(" ", "") == file_base.lower().replace(" ", ""):
                    png_file = v
                    break
        
        # 检查是否找到图片
        if png_file is None or not os.path.exists(png_file):
            print(f"找不到对应的背景图，跳过处理 {text_file}")
            print(f"  尝试查找: {district_name}")
            continue
        
        print(f"正在处理: {file_base} (使用背景图: {os.path.basename(png_file)})")
        
        # 读取文本文件
        with open(os.path.join(docs_folder, text_file), "r", encoding="utf-8") as f:
            text = f.read()
        
        # 分词
        words_list_jieba = jieba.lcut(text)
        
        # 构建词频字典
        word_dict = {}
        for key in words_list_jieba:
            word_dict[key] = word_dict.get(key, 0) + 1
        
        # 过滤停用词
        for demo in list(word_dict.keys()):
            if ('T' == findifhave(demo, stop)):
                del word_dict[demo]
        
        # 过滤掉长度不在2到4个字之间的词
        #filtered_dict = {key: value for key, value in word_dict.items() if 2 <= len(key) <= 4}
        
        # 过滤掉长度不在2到5个字之间的词
        filtered_dict = {}
        for key, value in word_dict.items():
            if 2 <= len(key) <= 5:
                # 增加3-5个字的词的权重
                if 3 <= len(key) <= 5:
                    filtered_dict[key] = int(value * 3.5)  # 增5加50%的权重
                else:  # 2个字的词保持原权重
                    filtered_dict[key] = value

        # 获取Top10高频词
        sorted_words = sorted(filtered_dict.items(), key=lambda x: x[1], reverse=True)
        top10_words = sorted_words[:10]
        
        # 保存区名、词频数据和图片路径
        all_districts_data[district_name] = {
            'top_words': top10_words,
            'all_words': filtered_dict,
            'png_file': png_file
        }
        
        # 生成并保存单个区的词云图
        generate_wordcloud(filtered_dict, district_name, png_file)
    
    return all_districts_data

# 生成词云图
def generate_wordcloud(word_dict, district_name, png_file):
    # 读取背景图片
    background_image = np.array(Image.open(png_file))
    
    # 设置颜色
    # 使用彩色系的颜色映射，类似于示例图
    colormaps = colors.ListedColormap(['#d53e4f', '#f46d43', '#fdae61', '#fee08b', '#e6f598', '#abdda4', '#66c2a5', '#3288bd'])
    
    # 生成词云
    wordcloud = WordCloud(font_path='simhei.ttf',
                          prefer_horizontal=0.9,
                          background_color='white',
                          max_words=100,
                          max_font_size=300,
                          stopwords=stop,
                          mask=background_image,
                        colormap=colormaps,
                          collocations=False
                         ).fit_words(word_dict)
    
    # 保存词云图到输出文件夹
    output_file = os.path.join(output_dir, f"{district_name}_词云图.png")
    wordcloud.to_file(output_file)
    print(f"  已保存词云图: {output_file}")
    
    return output_file

# 创建分组展示图
def create_group_visualization(group_idx, districts_data):
    group = district_groups[group_idx]
    group_name = f"第{group_idx+1}组"
    
    # 检查组内所有区是否都有数据
    valid_districts = [d for d in group if d in districts_data]
    if not valid_districts:
        print(f"警告: {group_name}没有有效的区域数据，跳过生成")
        return
    
    # 设置图表大小
    fig = plt.figure(figsize=(20, 14), constrained_layout=True)
    
    # 使用紧凑布局
    gs = fig.add_gridspec(nrows=2, ncols=1, height_ratios=[1, 1], hspace=0.05)
    
    # 添加总标题 - 高度接近图表顶部
    plt.suptitle(f"上海市十六区Top10高频特征词与词云图 - {group_name}", fontsize=22, y=0.98)
    
    # 创建表格数据
    table_data = []
    
    # 创建表头 - 每个区占两列：区名和词频
    header_row = []
    for d in valid_districts:
        header_row.extend([d, "词频"])
    
    # 准备表格数据
    for i in range(10):  # Top 10
        row = [f"Top{i+1}"]
        for district in valid_districts:
            if i < len(districts_data[district]['top_words']):
                word, freq = districts_data[district]['top_words'][i]
                row.extend([word, freq])
            else:
                row.extend(["", ""])
        table_data.append(row)
    
    # 创建表格区域
    ax_table = fig.add_subplot(gs[0])
    ax_table.axis('tight')
    ax_table.axis('off')
    
    # 创建表格，包括Top编号列
    the_table = ax_table.table(
        cellText=table_data,
        colLabels=[""] + header_row,  # 添加一个空单元格作为Top列的标题
        loc='center',
        cellLoc='center'
    )
    
    # 调整表格样式 - 增大字体，使表格纵向拉长
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(14)  # 增大字体大小
    the_table.scale(1, 2.0)     # 纵向拉长表格
    
    # 增加单元格边框宽度
    for (i, j), cell in the_table.get_celld().items():
        cell.set_linewidth(0.5)  # 设置边框宽度
        
        # 为表头设置粗体和背景色
        if i == 0:
            cell.set_text_props(weight='bold')
            cell.set_facecolor('#E6E6E6')  # 浅灰色背景
            
        # 对齐第一列（Top编号）
        if j == 0 and i > 0:
            cell.set_text_props(ha='center')
    
    # 创建词云图区域
    ax_wordcloud = fig.add_subplot(gs[1])
    ax_wordcloud.axis('off')
    
    # 在下半部分均匀排列词云图
    num_districts = len(valid_districts)
    
    # 为每个区添加词云图
    for i, district in enumerate(valid_districts):
        # 确定词云图的位置
        left = i / num_districts
        width = 1 / num_districts
        
        # 创建子图
        ax_wc = fig.add_axes([left, 0.05, width, 0.45])
        ax_wc.set_title(district, fontsize=16, pad=5)  # 减小标题和图的间距
        
        # 读取词云图
        img_path = os.path.join(output_dir, f"{district}_词云图.png")
        if os.path.exists(img_path):
            img = plt.imread(img_path)
            ax_wc.imshow(img)
            ax_wc.axis('off')
        else:
            print(f"警告: 找不到词云图 {img_path}")
    
    # 保存图表
    output_file = os.path.join(output_dir, f"上海市十六区_{group_name}_Top10高频特征词与词云图.png")
    plt.savefig(output_file, dpi=300, bbox_inches='tight', pad_inches=0.1)  # 减小边距
    plt.close()
    
    print(f"已生成分组可视化: {output_file}")

# 主函数
def main():
    print("开始生成上海市十六区Top10高频特征词与词云图...")
    
    # 处理所有区域的文本数据
    all_districts_data = process_all_districts()
    
    # 没有数据则结束
    if not all_districts_data:
        print("未找到任何区域数据，程序终止")
        return
    
    # 为每个分组生成可视化
    for i in range(len(district_groups)):
        create_group_visualization(i, all_districts_data)
    
    print("所有词频与词云图生成完毕！")

# 执行主函数
if __name__ == "__main__":
    main()