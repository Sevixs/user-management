#!/usr/bin/env python3
"""Fix Chinese curly quotes in the report generator."""
import sys

path = '/home/user/projects/user-management/generate_auth_report.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace Chinese left/right double quotation marks with corner brackets
content = content.replace('“', '「')  # " -> 「
content = content.replace('”', '」')  # " -> 」

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Done - fixed all Chinese curly quotes')
