#!/usr/bin/env python3
"""
调试真实的字符问题
"""

# 从错误信息中复制的实际字符串
error_line = '"instruction": "从目录中删除"摘要"项，保留其他内容和参考文献",'

print("错误行字符分析:")
print("=" * 50)
print(f"错误行: {repr(error_line)}")
print()

# 逐个字符分析
for i, char in enumerate(error_line):
    if ord(char) > 127 or char in ['"', '"', ''', ''']:
        print(f"位置 {i:2d}: {repr(char)} (U+{ord(char):04X}) - {char}")

print()
print("查找所有引号字符:")
quote_chars = []
for i, char in enumerate(error_line):
    if char in ['"', '"', '"', ''', ''']:
        quote_chars.append((i, char, ord(char)))
        print(f"位置 {i:2d}: {repr(char)} (U+{ord(char):04X})")

# 测试修复
print()
print("修复测试:")
fixed = error_line.replace('"', '"').replace('"', '"')
print(f"修复后: {repr(fixed)}")

# 测试JSON解析
import json
test_json = '{' + fixed + '"risk": "low"}'
print(f"测试JSON: {repr(test_json)}")

try:
    parsed = json.loads(test_json)
    print("✅ JSON解析成功!")
except json.JSONDecodeError as e:
    print(f"❌ JSON解析失败: {e}")
    print(f"错误位置: 第{e.lineno}行，第{e.colno}列")
    if e.colno <= len(test_json):
        print(f"错误字符: {repr(test_json[e.colno-1:e.colno+5])}")