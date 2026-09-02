#!/usr/bin/env bash
# =============================================================================
# make_skill_release.sh — 从 lesson-pack 源项目生成「合规的两级 WorkBuddy skill 包」
#
# 为什么需要它：
#   lesson-pack 目录同时承担两种角色：
#     1) 开源 git 项目（含 .git/.gitignore/LICENSE/README.md/assets 联系二维码）
#     2) WorkBuddy Skill 包（解析器要求：根/二级/文件，最多两级，不得含深层目录）
#   若直接把整个项目目录 cp 或打包，.git/objects/xx/... 的深层嵌套会触发
#   「目录层级超限 — Skill 包仅支持两级目录结构」。本脚本用 rsync 排除这些，
#   产出 skill 运行真正需要的干净两级副本与 .zip，可直接被解析/安装。
#
# 产出：
#   dist/lesson-pack/         干净的两级 skill 目录（无 .git/LICENSE/README/assets）
#   dist/lesson-pack.zip      上述目录的合规 zip
#
# 用法：
#   bash scripts/make_skill_release.sh
# =============================================================================
set -euo pipefail

# 定位到项目根（无论从哪调用）
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

DIST="$ROOT/dist"
OUT_DIR="$DIST/lesson-pack"
OUT_ZIP="$DIST/lesson-pack.zip"

echo "▶ 生成干净两级 skill 副本 → $OUT_DIR"
rm -rf "$OUT_DIR"
mkdir -p "$OUT_DIR"

# 只复制 skill 运行必需内容；排除 git/开源项目文件/生成物/系统垃圾
rsync -a \
  --exclude='.git/' \
  --exclude='.gitignore' \
  --exclude='LICENSE' \
  --exclude='README.md' \
  --exclude='assets/' \
  --exclude='scripts/make_skill_release.sh' \
  --exclude='dist/' \
  --exclude='data/' \
  --exclude='__pycache__/' \
  --exclude='*.py[cod]' \
  --exclude='*.egg-info/' \
  --exclude='.DS_Store' \
  ./ "$OUT_DIR/"

echo "▶ 打包 → $OUT_ZIP"
rm -f "$OUT_ZIP"
(cd "$DIST" && zip -r -X lesson-pack.zip lesson-pack -x '*.DS_Store' >/dev/null)

# 层级校验：Skill 包要求目录深度 ≤ 2（根/二级目录/文件），三级的目录即为违规
DEEP_COUNT="$(find "$OUT_DIR" -mindepth 3 -type d | wc -l | tr -d ' ')"
if [ "$DEEP_COUNT" -ne 0 ]; then
  echo "❌ 失败：发现 $DEEP_COUNT 个 3 级及以上目录，Skill 包不允许。请检查源项目是否多了深层子目录。"
  find "$OUT_DIR" -mindepth 3 -type d
  exit 1
fi

ZIP_DEEP="$(zipinfo -1 "$OUT_ZIP" 2>/dev/null | grep -cE 'lesson-pack/[^/]+/[^/]+/' || true)"
if [ "$ZIP_DEEP" -ne 0 ]; then
  echo "❌ 失败：zip 内存在 $ZIP_DEEP 条 3 级路径。"
  exit 1
fi

echo "✅ 完成。发布包已就绪，全部条目在两级以内，可直接被 skill 解析器安装。"
echo "   目录: $OUT_DIR"
echo "   zip : $OUT_ZIP"
echo
echo "   安装到用户级 Skills（可选）："
echo "   rsync -a \"$OUT_DIR/\" ~/.workbuddy/skills/lesson-pack/"
