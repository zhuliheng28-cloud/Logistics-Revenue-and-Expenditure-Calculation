#!/bin/bash

# 获取脚本所在目录
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# 定义 Python 脚本路径
PYTHON_SCRIPT="$DIR/income-expense-processor/scripts/process_excel.py"

echo "========================================"
echo "欢迎使用收支流水自动化处理工具"
echo "========================================"
echo ""
echo "请输入要处理的 Excel 文件路径 (可以直接将文件拖拽到此窗口中):"
read INPUT_FILE

# 去除输入路径可能包含的单引号（Mac终端拖拽文件时常见）
INPUT_FILE=$(echo "$INPUT_FILE" | sed "s/'//g" | xargs)

if [ ! -f "$INPUT_FILE" ]; then
    echo "错误：找不到文件 '$INPUT_FILE'"
    echo "按任意键退出..."
    read -n 1
    exit 1
fi

# 自动生成输出路径，在原文件名后加 _processed
OUTPUT_FILE="${INPUT_FILE%.*}_processed.xlsx"

echo ""
echo "正在处理，请稍候..."
echo "----------------------------------------"

# 运行 Python 脚本
python3 "$PYTHON_SCRIPT" "$INPUT_FILE" "$OUTPUT_FILE"

if [ $? -eq 0 ]; then
    echo "----------------------------------------"
    echo "✅ 处理成功！"
    echo "新文件已保存至: $OUTPUT_FILE"
    echo "您可以直接去该目录查看 Sheet2 的结果。"
else
    echo "----------------------------------------"
    echo "❌ 处理失败，请检查上面输出的错误信息。"
fi

echo ""
echo "按任意键退出..."
read -n 1
