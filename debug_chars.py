#!/usr/bin/env python3
"""
调试字符问题
"""

test_json = '''{"tasks": [{"id": "task_1","source_comment_id": "comment_1","type": "delete","locator": {"by": "find","value": "摘要"},"instruction": "从目录中删除"摘要"项，保留其他内容和参考文献","risk": "low"}]}'''

print("字符分析:")
print("=" * 50)

# 找到问题区域
problem_area = test_json[130:150]
print(f"问题区域: {repr(problem_area)}")

# 逐个字符分析
for i, char in enumerate(problem_area):
    print(f"{i:2d}: {repr(char)} (U+{ord(char):04X})")

print()
print("完整字符串的字符编码:")
chinese_quotes = ['"', '"', ''', ''']
for i, char in enumerate(test_json):
    if char in chinese_quotes:
        print(f"位置 {i}: {repr(char)} (U+{ord(char):04X})")