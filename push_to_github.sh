#!/bin/bash
# 一键上传到 GitHub
# 用法: ./push_to_github.sh YOUR_GITHUB_TOKEN

set -e

TOKEN="$1"
if [ -z "$TOKEN" ]; then
    echo "❌ 请提供 GitHub Personal Access Token"
    echo "用法: ./push_to_github.sh YOUR_GITHUB_TOKEN"
    echo ""
    echo "获取方式: GitHub → Settings → Developer settings → Personal access tokens"
    exit 1
fi

USERNAME="Dynamic-braking9"
REPO_NAME="chinese-government-docs"

echo "📦 创建 GitHub 仓库..."
curl -4 -sk -X POST \
  -H "Authorization: token $TOKEN" \
  https://api.github.com/user/repos \
  -d "{
    \"name\": \"$REPO_NAME\",
    \"description\": \"Hermes Agent Skill: 政企公文生成 (GB/T 9704-2012 标准)\",
    \"private\": false,
    \"auto_init\": false,
    \"license_template\": \"mit\"
  }" | python3 -c "
import sys, json
d = json.load(sys.stdin)
if 'full_name' in d:
    print(f'✅ 仓库已创建: {d[\"html_url\"]}')
elif 'already_exists' in json.dumps(d).lower():
    print('⚠️ 仓库已存在，继续推送...')
else:
    print(f'❌ 创建失败: {json.dumps(d, indent=2)}')
    sys.exit(1)
"

echo ""
echo "🚀 推送代码..."
cd "$(dirname "$0")"
git remote remove origin 2>/dev/null || true
git remote add origin "https://$USERNAME:$TOKEN@github.com/$USERNAME/$REPO_NAME.git"
git push -u origin main

echo ""
echo "🎉 完成！仓库地址: https://github.com/$USERNAME/$REPO_NAME"
