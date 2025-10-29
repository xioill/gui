#!/bin/zsh
# 快速初始化 git 并准备推送到 GitHub

set -euo pipefail

echo "正在初始化 Git 仓库..."

# 检查是否已经是 git 仓库
if [ -d .git ]; then
  echo "已存在 .git 目录，跳过初始化"
else
  git init
  echo "Git 仓库初始化完成"
fi

# 添加所有文件
git add .

# 检查是否有未提交的更改
if git diff --staged --quiet; then
  echo "没有需要提交的更改"
else
  git commit -m "初始提交：Excel 处理小工具"
  echo "已创建初始提交"
fi

echo ""
echo "接下来请执行："
echo "1. 在 GitHub.com 创建一个新仓库（https://github.com/new）"
echo "2. 复制仓库的 URL（例如：https://github.com/你的用户名/仓库名.git）"
echo "3. 运行以下命令推送代码："
echo ""
echo "   git remote add origin https://github.com/你的用户名/仓库名.git"
echo "   git branch -M main"
echo "   git push -u origin main"
echo ""

