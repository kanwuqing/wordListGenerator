
# [生词本生成器](https://github.com/kanwuqing/wordListGenerator)

将你的杂乱生词本自动整理成**可打印的表格文档**(.docx), 自动补全中文释义. 

## 功能

- 解析手写的原始生词文件(支持“单词 + 已有解释”和“仅单词”混合格式)并翻译成中文
- 同时保留你自己手动写的翻译(不会被覆盖)
- 生成一份排版整齐的 `.docx` 文件, 包含两列: 
  - **Word** – 单词 / 短语
  - **Meaning** – 中文释义
- 按字母排序

## 环境要求

- Python 3.6+
- pip

## 安装

### 1. 克隆项目(或直接下载脚本)

```bash
git clone https://github.com/kanwuqing/vocabulary-builder.git
cd vocabulary-builder
```

### 2. 安装 Python 依赖

```bash
python -m venv venv
pip install translate PyDictionary python-docx
```


## 使用方法

### 1. 准备你的生词文件 `raw.txt`

按照以下格式每行一条记录: 

- **只有单词**(脚本会自动翻译并添加词性)
  ```
  aberration
  give_up
  ```
  注意: 短语请用下划线连接, 如 `give_up`

- **单词 + 翻译**(用制表符或至少两个空格分隔, 翻译里可以包含词性)
  ```
  abandon	放弃
  abstract	adj. 抽象的
  ```
  词性部分(如 `adj.`)会被自动识别并提取到 Part of Speech 列, 剩下的文字作为释义. 

### 2. 运行脚本

```bash
python main.py raw.txt -o output.docx
```

终端会输出处理进度, 例如: 

```
已有翻译: 2 条, 未翻译: 2 条
  自动处理: aberration -> n.异常
  自动处理: give_up -> phrase放弃
✅ 词汇表已生成: 我的生词表.docx
```

## 文件说明

- `main.py` – 主脚本, 一键完成解析、翻译、生成表格
- `raw.txt` – 你手写的原始生词本(需要自行创建)
- `output.docx` – 最终生成的可打印词汇表(默认输出文件名)
- `-c 1|2|3` - 可选参数, 最终输出的表格的列数(默认为1列单词)

## 常见问题

**Q: 翻译失败或翻译结果不准确怎么办?**  
A: 翻译依赖网络和免费 API, 如果出现空释义, 可以稍后重新运行脚本. 你也可以直接在 `raw_words.txt` 里为部分单词手动写好翻译, 脚本会保留你的原始记录. 

**Q: 支持其他语言翻译吗?**  
A: 脚本默认是 **英文→中文**. 如需修改, 可编辑 `main.py` 中 `Translator(from_lang="en", to_lang="zh")` 的参数, 比如改成 `to_lang="es"`(西班牙语)等
