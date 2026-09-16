#!/usr/bin/env bash
# 回归：合成草稿 → 全流程 → 断言关键不变量
set -u
K="$(cd "$(dirname "$0")/.." && pwd)"
T="${1:-$(mktemp -d)}"
python3 "$K/tests/make_samples.py" "$T/samples" >/dev/null
fail=0
for s in s1_no_styles s2_cn_numbering s3_tracked; do
  P="$T/$s"; rm -rf "$P"
  python3 "$K/scripts/tfmt.py" init "$P" --draft "$T/samples/$s.docx" --profile tfsu-mti >/dev/null
  acc=$(python3 "$K/scripts/tfmt.py" accept --project "$P" 2>&1)
  echo "$acc" | grep -q '逐字一致' || { echo "✗ $s: accept 自证失败"; echo "$acc"; fail=1; continue; }
  out=$(python3 "$K/scripts/tfmt.py" outline --project "$P" 2>&1)
  n=$(echo "$out" | sed -n 's/.*判为标题 \([0-9]*\) 条.*/\1/p')
  sed -i.bak 's/# CONFIRMED: no/# CONFIRMED: yes/' "$P/02_outline.tsv"
  ap=$(python3 "$K/scripts/tfmt.py" apply --project "$P" 2>&1)
  echo "$ap" | grep -q '自证（草稿正文逐段原样保留）: ✓' || { echo "✗ $s: apply 自证失败"; echo "$ap"; fail=1; continue; }
  secs=$(python3 "$K/scripts/tfmt.py" audit --project "$P" --docx "$P/03_formatted.docx" 2>&1 | grep -c '^- 节')
  echo "✓ $s: 标题 $n 条, 分节 $secs 节"
  [ "$secs" -ge 2 ] || { echo "  ✗ 分节数异常"; fail=1; }
done
echo "---"; [ $fail -eq 0 ] && echo "回归全部通过（工作目录 $T）" || echo "有失败项（工作目录 $T）"
exit $fail
