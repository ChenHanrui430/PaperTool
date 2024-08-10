import os
import re
import pandas as pd
from PyPDF2 import PdfReader
from tqdm import tqdm
from collections import Counter
def count_keyword_frequency(keywords_data):
    """统计关键词的频率"""
    all_keywords = []
    for entry in keywords_data:
        all_keywords.extend(entry['keywords'].lower().split(', '))  # 无视大小写
    keyword_counts = Counter(all_keywords)
    return keyword_counts

def save_keywords_to_excel(keywords_data, keyword_counts, output_file):
    """将关键词数据和频率保存到 Excel 文件"""
    df = pd.DataFrame(keywords_data)
    df.to_excel(output_file, index=False, sheet_name='Keywords')

    # 统计结果另存为新表
    counts_df = pd.DataFrame(keyword_counts.items(), columns=['Keyword', 'Frequency'])
    counts_df.to_excel('freq'+output_file, index=False, sheet_name='Keyword Frequency')
def extract_keywords_from_pdf(pdf_path):
    """从 PDF 文件中提取关键词"""
    keywords = []
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            # 读取第一页的所有文本
            first_page_text = reader.pages[0].extract_text()
            # 使用正则表达式查找关键词部分，直到遇到 DOI、空行或非单词字符
            match = re.search(r'Keywords?:?\s*(.*?)(?=\n\s*DOI|\n\s*$|\n\W|\n\s*This|\n\s*The|\n\s*1|\n\s*ACKNOWLEDGMENTS)', first_page_text, re.DOTALL | re.IGNORECASE)
            if match:
                if len(match.group(1))<200:
                # 分割关键词并清理空白
                    keywords = [kw.strip() for kw in match.group(1).split(',')]
            else:
                pass
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return keywords

def collect_keywords_from_folder(folder_path):
    """从指定文件夹中的所有 PDF 文件收集关键词"""
    all_keywords = []
    for filename in tqdm(os.listdir(folder_path)):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            keywords = extract_keywords_from_pdf(pdf_path)
            all_keywords.append({'filename': filename, 'keywords': ', '.join(keywords)})
    return all_keywords

def main():
    folder_path = 'E:\download\popets'  # 替换为您的文件夹路径
    output_file = 'keywords.xlsx'  # 输出的 Excel 文件名

    keywords_data = collect_keywords_from_folder(folder_path)
    keyword_counts = count_keyword_frequency(keywords_data)
    save_keywords_to_excel(keywords_data, keyword_counts, output_file)
    print(f"Keywords have been saved to {output_file}")

if __name__ == "__main__":
    main()