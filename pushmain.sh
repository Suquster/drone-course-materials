#!/bin/bash
# Push current HEAD to the GitHub remote for this repo.
# Usage: ./pushmain.sh [branch=content]
# Auth: export GitHubPAT=<token>
set -e
TARGET_BRANCH="${1:-content}"
if [ -z "${GitHubPAT:-}" ]; then
  echo "错误: 未设置 GitHubPAT 环境变量" >&2
  exit 1
fi
git push "https://${GitHubPAT}@github.com/Suquster/drone-course-materials.git" "HEAD:refs/heads/${TARGET_BRANCH}"
echo "Pushed to ${TARGET_BRANCH} successfully"
