#!/usr/bin/env bash
# NutraPLM 一键启动：检查/安装依赖 → 启动后端（同时提供前端静态资源）
set -e
cd "$(dirname "$0")"

PY=${PYTHON:-python3}

check_deps() {
  $PY - <<'PY'
import importlib.util, sys
need = {"fastapi": "fastapi", "uvicorn": "uvicorn", "openpyxl": "openpyxl", "multipart": "python-multipart"}
missing = [pkg for mod, pkg in need.items() if not importlib.util.find_spec(mod)]
if missing:
    print("缺少依赖：" + ", ".join(missing))
sys.exit(1 if missing else 0)
PY
}

echo "==> 检查依赖"
if ! check_deps; then
  echo "==> 安装依赖（pip install -r server/requirements.txt）"
  $PY -m pip install -q -r server/requirements.txt
fi

echo "==> 启动服务： http://127.0.0.1:${PORT:-8000}"
echo "    首次启动会自动建库并写入演示数据（server/plm.db）"
echo "    演示账号：admin / admin123（其他角色见 README）"
exec $PY -m uvicorn app:app --app-dir server --host 0.0.0.0 --port "${PORT:-8000}"
