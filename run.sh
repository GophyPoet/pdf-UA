#!/usr/bin/env bash
# pdf/UA launcher — Linux / macOS.
# Двойной клик из файлового менеджера или ./run.sh из терминала.
# Стартует локальный Flask-сервер и открывает UI в браузере.

set -e
cd "$(dirname "$0")"

PY=${PYTHON:-python3}

# Проверка зависимостей: тихо доустанавливаем pip-пакеты,
# системные (libreoffice, tesseract) — на пользователе.
if ! "$PY" -c "import flask, PIL, pytesseract" >/dev/null 2>&1; then
  echo "[pdf/UA] устанавливаю Python-зависимости…"
  "$PY" -m pip install --quiet -r requirements.txt
fi

if ! command -v soffice >/dev/null 2>&1 && ! command -v libreoffice >/dev/null 2>&1; then
  echo "[pdf/UA] ОШИБКА: не найден LibreOffice (soffice). Установите libreoffice."
  exit 1
fi
if ! command -v tesseract >/dev/null 2>&1; then
  echo "[pdf/UA] ПРЕДУПРЕЖДЕНИЕ: не найден tesseract — OCR работать не будет."
fi

PORT=${PDFUA_PORT:-8000}
HOST=${PDFUA_HOST:-127.0.0.1}

echo "[pdf/UA] запуск на http://$HOST:$PORT/"
exec "$PY" -m pdfua.cli serve --host "$HOST" --port "$PORT" --open
