import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import matplotlib
import os
import re
from collections import Counter

# 设置中文字体支持
matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
matplotlib.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 确保输出文件夹存在
output_dir = '词频与词云图'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 辅助函数：检查是否为停用词
def findifhave(demo, stop):
    for ret in stop:
        if (demo == ret):
            return 'T'

# 获取停用词列表
def load_stopwords(stopwords_file="stop.txt"):
    stop = ['\n']
    try:
        with open(stopwords_file, 'r', encoding='utf-8') as f1:
            for line in f1:
                stop.append(line.replace("\n", ""))
        return stop
    except FileNotFoundError:
        print(f"警告: 停用词文件 '{stopwords_file}' 未找到，使用默认停用词列表")
        return stop

# 从文件名中提取区域名称
def extract_district_name(filename):
    # 尝试匹配"XX区"或"XX新区"格式
    match = re.search(r'([^\s]+[区]|[^\s]+新区)', filename)
    if match:
        return match.group(0)
    return os.path.splitext(filename)[0].strip()  # 如果没有匹配到，返回无扩展名的文件名

# 处理文本并获取词频
def process_text(text_file, input_folder, stop_words):
    district_name = extract_district_name(text_file)
    
    print(f"正在处理: {district_name}")
    
    # 读取文本文件
    with open(os.path.join(input_folder, text_file), "r", encoding="utf-8") as f:
        text = f.read()
    
    # 分词
    words_list_jieba = jieba.lcut(text)
    
    # 构建词频字典
    word_dict = {}
    for key in words_list_jieba:
        word_dict[key] = word_dict.get(key, 0) + 1
    
    # 过滤停用词
    for demo in list(word_dict.keys()):
        if ('T' == findifhave(demo, stop_words)):
            del word_dict[demo]
    
    # 仅保留2-5字词汇（不调整权重）
    filtered_dict = {k: v for k, v in word_dict.items() if 2 <= len(k) <= 5}
    
    # 计算总词频
    total = sum(filtered_dict.values())
    
    # 获取Top50高频词并计算占比
    sorted_words = sorted(filtered_dict.items(), key=lambda x: x[1], reverse=True)
    top50_words = [(word, freq, f"{freq/total*100:.2f}%") for word, freq in sorted_words[:50]]
    
    return district_name, filtered_dict, top50_words, total

# 生成词云图
def generate_wordcloud(word_dict, title):
    # 生成词云
    wordcloud = WordCloud(
        font_path='simhei.ttf',
        width=800,
        height=600,
        background_color='black',  # 修改为黑色背景
        max_words=100,
        max_font_size=300,
        prefer_horizontal=0.9,
        color_func=lambda *args, **kwargs: "white",  # 设置文字为白色
        collocations=False
    ).fit_words(word_dict)
    
    # 创建图形并显示词云
    plt.figure(figsize=(10, 8))
    plt.imshow(wordcloud, interpolation="bilinear")
    plt.axis("off")
    # 移除标题行
    
    # 保存词云图
    output_file = os.path.join(output_dir, f"{title}_词云图.png")
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"  已保存词云图: {output_file}")
    return output_file

# # 创建词频表格
# def create_word_frequency_table(district_name, top_words):
#     # 创建表格
#     fig, ax = plt.figure(figsize=(8, 6)), plt.gca()
#     ax.axis('tight')
#     ax.axis('off')
    
#     # 准备表格数据
#     table_data = [[f"Top{i+1}", word, freq] for i, (word, freq) in enumerate(top_words)]
    
#     # 创建表格
#     table = plt.table(
#         cellText=table_data,
#         colLabels=["排名", "词语", "词频"],
#         loc='center',
#         cellLoc='center'
#     )
    
#     # 调整表格样式
#     table.auto_set_font_size(False)
#     table.set_fontsize(12)
#     table.scale(1, 1.5)
    
#     # 移除标题
    
#     # 保存表格
#     output_file = os.path.join(output_dir, f"{district_name}_词频表.png")
#     plt.savefig(output_file, dpi=300, bbox_inches='tight')
#     plt.close()
    
#     print(f"  已保存词频表: {output_file}")
#     return output_file

# 创建完整词频表格
def create_full_frequency_table(district_name, top50_data):
    # 创建表格数据
    table_data = []
    for i in range(25):  # 两列各25行
        row = [
            i+1,
            top50_data[i][0],
            f"{top50_data[i][1]:,}".replace(",", " "),
            top50_data[i][2],
            i+26,
            top50_data[i+25][0],
            f"{top50_data[i+25][1]:,}".replace(",", " "),
            top50_data[i+25][2]
        ]
        table_data.append(row)

    # 创建图表
    fig, ax = plt.subplots(figsize=(16, 12))
    ax.axis('off')
    
    # 创建表格
    table = ax.table(
        cellText=table_data,
        colLabels=["序号", "主题词", "词频 / 次", "占比", "序号", "主题词", "词频 / 次", "占比"],
        colColours=['#f0f0f0']*8,
        cellLoc='center',
        loc='center'
    )
    
    # 设置表格样式
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 2)  # 调整表格尺寸
    
    # 设置标题
    plt.title(f"{district_name} xxx词频排名前50位", 
             fontsize=14, y=0.95, pad=20)
    
    # 保存图片
    output_file = os.path.join(output_dir, f"{district_name}_词频总表.png")
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"  已保存完整词频表: {output_file}")
    return output_file

# 主函数
def main():
    # 设置文档文件夹路径
    docs_folder = "xxx"  # 修改为实际的文档文件夹路径
    
    print("开始生成词频统计与词云图...")
    
    # 加载停用词
    stop_words = load_stopwords()
    
    # 获取所有文本文件
    if not os.path.exists(docs_folder):
        print(f"错误: 找不到文档文件夹 '{docs_folder}'")
        return
    
    text_files = [f for f in os.listdir(docs_folder) 
                 if os.path.isfile(os.path.join(docs_folder, f)) 
                 and f.endswith('.txt')]
    
    if not text_files:
        print(f"错误: 在 '{docs_folder}' 中没有找到文本文件")
        return
    
    # 处理每个文本文件
    for text_file in text_files:
        # 修改返回值接收
        district_name, word_dict, top50_words, total = process_text(text_file, docs_folder, stop_words)
        
        # 生成完整词频表
        create_full_frequency_table(district_name, top50_words)
        
        # 生成词云图
        generate_wordcloud(word_dict, district_name)
        
        # # 生成词频表格 (使用Top10)
        # create_word_frequency_table(district_name, top50_words[:10])
    
    print("所有词频统计与词云图生成完毕！")

# 执行主函数
if __name__ == "__main__":
    main()