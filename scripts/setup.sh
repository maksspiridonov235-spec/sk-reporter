#!/usr/bin/env bash
# Первичная настройка SK-Reporter (macOS/Linux). Запуск из корня репозитория.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if [[ ! -f webapp/main.py ]]; then
  echo "Запустите скрипт из клона sk-reporter (нужен webapp/main.py)." >&2
  exit 1
fi

if [[ -d webapp/venv ]]; then
  echo "Удаляю устаревший webapp/venv..."
  rm -rf webapp/venv
fi

echo "Создаю venv в корне..."
python3 -m venv venv
# shellcheck source=/dev/null
source venv/bin/activate
pip install --upgrade pip
pip install -e .

echo ""
echo "Готово. Запуск:"
echo "  source venv/bin/activate"
echo "  cd webapp"
echo "  python3 -m uvicorn main:app --reload --host 127.0.0.1 --port 8000"
echo ""
