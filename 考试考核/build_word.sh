#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

OUT_MD="_merged.md"
OUT_DOCX="《Web项目实战》课程考试考核材料汇编.docx"

# 1. 生成合并后的 Markdown（带封面、目录、章节）
{
  echo '---'
  echo 'title: "《Web 项目实战》课程考试考核材料汇编"'
  echo 'subtitle: "课程代码：32093143　　编制依据：课程目标达成度评价报告 v2.0"'
  echo 'author: "课程负责人：廖清科"'
  echo 'date: "重庆工程职业技术学院　大数据与物联网学院"'
  echo 'lang: zh-CN'
  echo '---'
  echo
  echo '# 封面与说明'
  echo
  echo '| 项目 | 内容 |'
  echo '| --- | --- |'
  echo '| 课程名称 | Web 项目实战 |'
  echo '| 课程代码 | 32093143 |'
  echo '| 开课院部 | 大数据与物联网学院 |'
  echo '| 适用专业 | 软件技术 |'
  echo '| 考核性质 | 考试 |'
  echo '| 课程性质 | 专业核心课（理论+实践） |'
  echo '| 总学时 | 64（理论 32 + 实践 32） |'
  echo '| 编制依据 | 《Web 项目实战》课程目标达成度评价报告 v2.0 |'
  echo '| 编制单位 | 重庆工程职业技术学院　大数据与物联网学院 |'
  echo '| 编制人 | 廖清科 |'
  echo '| 编制时间 | 2026 年 4 月 |'
  echo
  echo '本汇编收录《Web 项目实战》课程的完整考试/考核材料，包含考核总方案、过程性考核材料、结果性考核材料以及命题审批与评分表四大部分，覆盖 T1~T10 全部课程目标，支撑课程目标达成度评价工作。'
  echo
  echo '\newpage'
  echo

  # 章节映射：目录 -> 大章节标题
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
      fname="$(basename "$f" .md)"
      echo
      echo "## $fname"
      echo
      # 将原文件中的一级标题降级为三级，二级降为四级，依次类推
      sed -E 's/^###### /######## /; s/^##### /####### /; s/^#### /###### /; s/^### /##### /; s/^## /#### /; s/^# /### /' "$f"
      echo
      echo '\newpage'
      echo
    done
  done
} > "$OUT_MD"

echo "已生成合并 Markdown: $OUT_MD ($(wc -l < "$OUT_MD") 行)"

# 2. 使用 pandoc 转换为 docx
pandoc "$OUT_MD" \
  -f markdown+pipe_tables+yaml_metadata_block \
  -t docx \
  --toc \
  --toc-depth=2 \
  --number-sections \
  -M toc-title="目　　录" \
  -o "$OUT_DOCX"

echo "已生成 Word 文档: $OUT_DOCX"
ls -lh "$OUT_DOCX"

rm -f "$OUT_MD"
