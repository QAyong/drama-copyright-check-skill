"""
短剧剧本预处理脚本（格式转换 + 分段）
支持 .txt / .doc / .docx / .pdf 输入，超过 8 万字自动切分。

用法：
    python prepare.py <剧本文件>
    python prepare.py 我的剧本.docx
    python prepare.py 对比剧本.pdf

依赖（按需安装）：
    pip install python-docx pdfplumber pywin32
"""

import re
import sys
from pathlib import Path

THRESHOLD = 80000
CHUNK_SIZE = 50000

EPISODE_PATTERN = re.compile(
    r'(?=(?:第\s*\d+\s*集|EP\.?\s*\d+|第\s*[一二三四五六七八九十百]+\s*集))',
    re.IGNORECASE
)


# ---------- 格式读取 ----------

def read_txt(path: Path) -> str:
    for enc in ('utf-8', 'gbk', 'utf-16'):
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, ValueError):
            continue
    raise ValueError(f"无法识别文件编码：{path}")


def read_docx(path: Path) -> str:
    try:
        from docx import Document
    except ImportError:
        print("缺少依赖，请先运行：pip install python-docx")
        sys.exit(1)
    doc = Document(str(path))
    return '\n'.join(p.text for p in doc.paragraphs)


def read_doc(path: Path) -> str:
    try:
        import win32com.client
    except ImportError:
        print("缺少依赖，请先运行：pip install pywin32")
        sys.exit(1)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(path.resolve()))
        text = doc.Content.Text
        doc.Close(False)
    finally:
        word.Quit()
    return text


def read_pdf(path: Path) -> str:
    try:
        import pdfplumber
    except ImportError:
        print("缺少依赖，请先运行：pip install pdfplumber")
        sys.exit(1)
    pages = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages.append(text)
    return '\n'.join(pages)


def load_file(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == '.txt':
        return read_txt(path)
    elif ext == '.docx':
        return read_docx(path)
    elif ext == '.pdf':
        return read_pdf(path)
    elif ext == '.doc':
        return read_doc(path)
    else:
        print(f"不支持的文件格式：{ext}（支持 .txt / .docx / .pdf）")
        sys.exit(1)


# ---------- 分段逻辑 ----------

def split_by_episodes(text: str) -> list[str]:
    parts = EPISODE_PATTERN.split(text)
    return [p.strip() for p in parts if p.strip()]


def split_by_size(text: str, chunk_size: int) -> list[str]:
    return [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]


def merge_small_parts(parts: list[str], min_size: int = 5000) -> list[str]:
    merged = []
    for part in parts:
        if merged and len(part) < min_size:
            merged[-1] += '\n\n' + part
        else:
            merged.append(part)
    return merged


# ---------- 主流程 ----------

def main():
    if len(sys.argv) < 2:
        print("用法：python prepare.py <剧本文件路径>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"文件不存在：{input_path}")
        sys.exit(1)

    # 1. 读取并转换为纯文本
    print(f"读取文件：{input_path.name}（格式：{input_path.suffix}）")
    text = load_file(input_path)
    char_count = len(text)
    print(f"提取字符数：{char_count:,}")

    # 2. 保存中间纯文本（方便排查问题）
    out_dir = input_path.parent / (input_path.stem + '_parts')
    out_dir.mkdir(exist_ok=True)
    plain_file = out_dir / 'full_text.txt'
    plain_file.write_text(text, encoding='utf-8')
    print(f"纯文本已保存：{plain_file}")

    # 3. 长度判断
    if char_count <= THRESHOLD:
        print(f"\n字符数未超过 {THRESHOLD:,}，无需分段。")
        print(f"直接将 full_text.txt 内容传入 drama-copyright-check skill 即可。")
        return

    # 4. 分段
    parts = split_by_episodes(text)
    if len(parts) > 1:
        parts = merge_small_parts(parts)
        mode = "按分集标记切分"
    else:
        parts = split_by_size(text, CHUNK_SIZE)
        mode = f"无分集标记，按 {CHUNK_SIZE:,} 字符均分"

    for i, part in enumerate(parts, 1):
        (out_dir / f'part_{i}.txt').write_text(part, encoding='utf-8')

    # 5. 操作指引
    print(f"\n切分模式：{mode}")
    print(f"共 {len(parts)} 段，输出目录：{out_dir}\n")
    print("=" * 52)
    print("操作步骤：")
    print("1. 按顺序将以下各段依次传入 drama-copyright-check skill：")
    for i, part in enumerate(parts, 1):
        print(f"   第 {i} 段：part_{i}.txt（{len(part):,} 字符）")
    print("2. 每段处理完后继续传入下一段，无需重复说明任务。")
    print("3. 全部段落传完后说：「以上是全部分段，请合并输出完整对照表。」")
    print("=" * 52)


if __name__ == '__main__':
    main()
