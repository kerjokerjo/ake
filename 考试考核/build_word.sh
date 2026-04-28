#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

OUT_MD="_merged.md"
OUT_DOCX="《Web项目实战》课程考试考核材料汇编.docx"

# 层级规则：
#   一级 (#)  —— 第X篇
#   二级 (##) —— 每份 md 文件的标题
#   文件内部原为 "# 标题"，合并时统一降为 "## 标题"；文件内不再使用更深层级
# 合并后的最大标题层级为 2，满足"最多三级"要求

{
  declare -A SECTIONS=(
    ["01_考核总方案"]="第一篇 考核总方案"
    ["02_过程性考核"]="第二篇 过程性考核"
    ["03_结果性考核"]="第三篇 结果性考核"
    ["04_命题审批与评分表"]="第四篇 命题审批与评分表"
  )

  for dir in "01_考核总方案" "02_过程性考核" "03_结果性考核" "04_命题审批与评分表"; do
    echo
    echo "# ${SECTIONS[$dir]}"
    echo
    for f in "$dir"/*.md; do
      [[ "$(basename "$f")" == "README.md" ]] && continue
      echo
      # 将文件中的一级标题降为二级；文件内部目前没有更深 heading
      sed -E 's/^# /## /' "$f"
      echo
      echo '\newpage'
      echo
    done
  done
} > "$OUT_MD"

echo "已生成合并 Markdown: $OUT_MD ($(wc -l < "$OUT_MD") 行)"

pandoc "$OUT_MD" \
  -f markdown+pipe_tables \
  -t docx \
  --toc \
  --toc-depth=2 \
  -M toc-title="目　　录" \
  -o "$OUT_DOCX"

echo "已生成 Word 文档: $OUT_DOCX"
ls -lh "$OUT_DOCX"

rm -f "$OUT_MD"
