#!/usr/bin/env python3
"""
自动翻译生词，合并重复条目，生成 docx 表格，并输出翻译后的词汇文本文件。
用法: python build_vocab.py raw_words.txt [-o output.docx]
"""

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
    """直接翻译单词或短语，返回中文释义"""
    display_word = raw_word.replace('_', ' ')
    try:
        meaning = translator.translate(display_word)
    except:
        meaning = ''
    return meaning if meaning else ''

def parse_and_merge(filepath):
    """
    解析词汇文件，自动合并重复单词。
    规则：若有任意一行包含有效释义，则保留；若全都没有释义，则记为 None。
    返回 OrderedDict: {word: meaning_or_None}
    """
    merged = OrderedDict()
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = re.split(r'\t| {2,}', line, maxsplit=1)
            word = parts[0].strip()
            meaning = None
            if len(parts) > 1:
                rest = parts[1].strip()
                # 去除可能存在的词性标记（如 "n. 放弃" → 只保留释义）
                pos_match = re.match(r'\(?([a-z]+\.)\)?\s*(.*)', rest, re.IGNORECASE)
                if pos_match:
                    meaning = pos_match.group(2).strip()
                else:
                    meaning = rest
                # 如果提取出的 meaning 是空字符串，视为无释义
                if not meaning:
                    meaning = None

            # 合并逻辑：有释义则更新，无释义且单词未出现过才记为 None
            if meaning:
                merged[word] = meaning
            elif word not in merged:
                merged[word] = None
    return merged

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('raw_file')
    parser.add_argument('-o', '--output', default='vocabulary.docx')
    args = parser.parse_args()

    # 合并与解析
    vocab = parse_and_merge(args.raw_file)
    unknown_count = sum(1 for v in vocab.values() if not v)
    known_count = len(vocab) - unknown_count
    print(f"已有翻译: {known_count} 条，未翻译: {unknown_count} 条")

    # 翻译所有无释义的单词
    for word in list(vocab.keys()):
        if not vocab[word]:
            meaning = auto_translate_word(word)
            vocab[word] = meaning
            print(f"\t自动处理: {word} -> {meaning}")

    # 排序（保持输入顺序或字母顺序，这里用字母）
    sorted_words = sorted(vocab.keys(), key=lambda w: w.replace('_', ' ').lower())

    # ---------- 生成 docx ----------
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    title = doc.add_heading(f'Vocabulary List ({os.path.basename(args.raw_file)})', level=1)
    title.alignment = WD_TABLE_ALIGNMENT.CENTER

    table = doc.add_table(rows=1, cols=2, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Word'
    hdr_cells[1].text = 'Meaning'
    for i in range(2):
        for p in hdr_cells[i].paragraphs:
            p.style.font.bold = True

    for word in sorted_words:
        row_cells = table.add_row().cells
        row_cells[0].text = word.replace('_', ' ')
        row_cells[1].text = vocab[word] if vocab[word] else ''

    # 调整列宽
    for cell in table.columns[0].cells:
        cell.width = Inches(2.0)
    for cell in table.columns[1].cells:
        cell.width = Inches(4.5)

    doc.save(args.output)
    print(f"✅ 词汇表已生成: {args.output}")

    # ---------- 输出回文本文件 ----------
    base, ext = os.path.splitext(args.raw_file)
    out_txt = f"{base}-output{ext}"
    with open(out_txt, 'w', encoding='utf-8') as f:
        for word in sorted_words:
            meaning = vocab[word]
            if meaning:
                f.write(f"{word}\t{meaning}\n")
            else:
                f.write(f"{word}\n")
    print(f"📄 完整词汇表已保存: {out_txt}")

if __name__ == '__main__':
    main()