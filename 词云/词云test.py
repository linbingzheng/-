import jieba
from wordcloud import WordCloud
import numpy as np
from PIL import Image
from matplotlib import colors
import os
import re

# 确保输出文件夹存在
output_dir = '所有词云图'
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

# 打印可用的图片文件，帮助调试
print("可用的背景图片文件:")
if os.path.exists(images_folder):
    png_files = [f for f in os.listdir(images_folder) if f.lower().endswith('.png')]
    for png in png_files:
        print(f"  - {png}")
else:
    print(f"警告: 找不到图片文件夹 '{images_folder}'")

# 提取区名的函数 - 只保留"XX区"或"XX新区"部分
def extract_district_name(filename):
    # 尝试匹配"XX区"或"XX新区"格式
    match = re.search(r'([^\s]+[区]|[^\s]+新区)', filename)
    if match:
        return match.group(0)
    return os.path.splitext(filename)[0].strip()  # 如果没有匹配到，返回无扩展名的文件名

# 获取所有文本文件
text_files = []
if os.path.exists(docs_folder):
    text_files = [f for f in os.listdir(docs_folder) 
                 if os.path.isfile(os.path.join(docs_folder, f)) 
                 and (f.endswith('.txt') or f.endswith('.xlsx'))]
else:
    print(f"警告: 找不到文档文件夹 '{docs_folder}'")

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

for text_file in text_files:
    # 获取文件名（不含扩展名）
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
    
    # 对于Excel文件，我们需要不同的处理方式，这里简化为跳过
    if text_file.lower().endswith('.xlsx'):
        print(f"  跳过Excel文件: {text_file} - 请先转换为文本格式")
        continue
    
    # 读取文本文件
    with open(os.path.join(docs_folder, text_file), "r", encoding="utf-8") as f:
        text = f.read()
    
    # 分词
    words_list_jieba = jieba.lcut(text)
    
    # 构建词频字典
    dict = {}
    for key in words_list_jieba:
        dict[key] = dict.get(key, 0) + 1
    
    # 过滤停用词
    for demo in list(dict.keys()):
        if ('T' == findifhave(demo, stop)):
            del dict[demo]
    
    # 过滤掉长度不在2到4个字之间的词
    # filtered_dict = {key: value for key, value in dict.items() if 2 <= len(key) <= 4}
    filtered_dict = {}
    for key, value in dict.items():
        if 3 <= len(key) <= 5:
            # 增强3-5个字词语的权重
            filtered_dict[key] = value * 3.5
        elif len(key) == 2:
            # 减弱2个字词语的权重，作为辅助
            filtered_dict[key] = value * 0.5
    # 读取背景图片
    background_image = np.array(Image.open(png_file))
    
    # 设置颜色
    colormaps = colors.ListedColormap(['#871A84', '#BC0F6A', '#BC0F60', '#CC5F6A', '#AC1F4A'])
    
    # 生成词云
    wordcloud = WordCloud(font_path='simhei.ttf',
                          prefer_horizontal=0.99,
                          background_color='white',
                          max_words=90,
                          max_font_size=400,
                          stopwords=stop,
                          mask=background_image,
                          colormap=colormaps,
                          collocations=False
                         ).fit_words(filtered_dict)
    
    # 保存词云图到输出文件夹
    output_file = os.path.join(output_dir, f"{district_name}_词云图.png")
    wordcloud.to_file(output_file)
    print(f"  已保存词云图: {output_file}")

print("所有词云图生成完毕！")
