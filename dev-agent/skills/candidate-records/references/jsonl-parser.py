import re
import json
import sys

# 用法: python jsonl-parser.py <输入文件> [输出文件]
if len(sys.argv) >= 3:
    input_file = sys.argv[1]
    output_file = sys.argv[2]
elif len(sys.argv) == 2:
    input_file = sys.argv[1]
    output_file = "output.jsonl"
else:
    input_file = input("请输入面试文本文件路径: ").strip()
    output_file = input("请输入输出JSONL文件路径: ").strip()

# 读取文本文件
with open(input_file, "r", encoding="utf-8") as f:
    text = f.read()

# 按行读取，每两行一组（说话内容 + 说话人时间）
lines = text.strip().split('\n')
results = []

i = 0
while i < len(lines):
    content = lines[i].strip()
    # 检查下一行是否是时间戳行（包含空格分隔的姓名和时间）
    if i + 1 < len(lines):
        time_line = lines[i + 1].strip()
        # 匹配格式: " 姓名 HH:MM:SS" 或 "姓名 HH:MM:SS"
        match = re.match(r'^\s*(\S+)\s+(\d{2}:\d{2}:\d{2})$', time_line)
        if match:
            name = match.group(1)
            timestamp = match.group(2)
            results.append({
                "姓名": name,
                "时间戳": timestamp,
                "内容": content
            })
            i += 2
            continue
    # 无法匹配的行，作为内容处理
    results.append({
        "姓名": "",
        "时间戳": "",
        "内容": content
    })
    i += 1

# 直接写入 JSONL
with open(output_file, "w", encoding="utf-8") as f:
    for line in results:
        f.write(json.dumps(line, ensure_ascii=False) + "\n")

print(f"Conversion done! Output: {output_file}")