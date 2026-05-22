#!/usr/bin/env python3

import re
import os
import argparse
from collections import OrderedDict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from translate import Translator

translator = Translator(from_lang="en", to_lang="zh")

def auto_translate_word(raw_word):
    """直接翻译单词或短语,返回中文释义"""
    display_word = raw_word.replace('_', ' ')
    try:
        meaning = translator.translate(display_word)
    except:
        meaning = ''
    return meaning if meaning else ''

def parse_raw_file(filepath):
    known = OrderedDict()
    unknown = OrderedDict() 

    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = re.split(r'\t| {2,}', line, maxsplit=1)
            word = parts[0].strip()
            if len(parts) == 1:
                # 仅有单词
                if word not in unknown:
                    unknown[word] = None
            else:
                rest = parts[1].strip()
                # 去除可能存在的词性标记（如 "n. 放弃" → 只保留释义）
                pos_match = re.match(r'\(?([a-z]+\.)\)?\s*(.*)', rest, re.IGNORECASE)
                if pos_match:
                    # 词性部分忽略，只取后面的释义
                    meaning = pos_match.group(2).strip()
                else:
                    meaning = rest
                if word not in known:
                    known[word] = meaning if meaning else None

    return known, unknown

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('raw_file')
    parser.add_argument('-o', '--output', default='vocabulary.docx')
    args = parser.parse_args()

    known, unknown = parse_raw_file(args.raw_file)
    print(f"已有翻译: {len(known)} 条，未翻译: {len(unknown)} 条")

    # 翻译所有未翻译单词
    for word in sorted(unknown.keys()):
        meaning = auto_translate_word(word)
        known[word] = meaning
        print(f"\t自动处理: {word} -> {meaning}")

    # 按字母排序
    sorted_words = sorted(known.keys(), key=lambda w: w.replace('_', ' ').lower())

    # 生成 docx
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    title = doc.add_heading(f'Vocabulary List({os.path.basename(args.raw_file)})', level=1)
    title.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 只保留两列：Word, Meaning
    table = doc.add_table(rows=1, cols=2, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Word'
    hdr_cells[1].text = 'Meaning'
    for i in range(2):
        for p in hdr_cells[i].paragraphs:
            p.style.font.bold = False

    for word in sorted_words:
        meaning = known[word]
        row_cells = table.add_row().cells
        row_cells[0].text = word.replace('_', ' ')
        row_cells[1].text = meaning if meaning else ''

    # 调整列宽
    for cell in table.columns[0].cells:
        cell.width = Inches(2.0)
    for cell in table.columns[1].cells:
        cell.width = Inches(4.5)

    doc.save(args.output)
    print(f"✅ 词汇表已生成: {args.output}")

if __name__ == '__main__':
    main()