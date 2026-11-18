#!/usr/bin/env bash
# Codex CLI に非対話でタスクを渡し、最終回答だけをファイルと標準出力に返す。
# 使い方:
#   ask_codex.sh [-w] [-m MODEL] [-C DIR] [-o OUTFILE] [--resume] < prompt.txt
#   -w        ワークスペース書き込みを許可（既定は read-only）
#   -m        モデル指定（未指定なら Codex の既定）
#   -C        作業ディレクトリ（既定はカレント）
#   -o        最終回答の保存先（既定 .codex-out/<timestamp>.md）
#   --resume  直前の Codex セッションの続きとして送る
set -euo pipefail

SANDBOX="read-only"; MODEL=""; DIR="$PWD"; OUT=""; RESUME=0
while [[ $# -gt 0 ]]; do
  case "$1" in
    -w) SANDBOX="workspace-write"; shift ;;
    -m) MODEL="$2"; shift 2 ;;
    -C) DIR="$2"; shift 2 ;;
    -o) OUT="$2"; shift 2 ;;
    --resume) RESUME=1; shift ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

if ! command -v codex >/dev/null 2>&1; then
  echo "ERROR: codex CLI が見つかりません。'npm i -g @openai/codex' でインストールし、'codex login' で認証してください。" >&2
  exit 127
fi

PROMPT="$(cat)"
[[ -z "$PROMPT" ]] && { echo "ERROR: プロンプトが空です（標準入力で渡してください）" >&2; exit 2; }

mkdir -p "$DIR/.codex-out"
OUT="${OUT:-$DIR/.codex-out/$(date +%Y%m%d-%H%M%S).md}"

ARGS=(exec --skip-git-repo-check -C "$DIR" -s "$SANDBOX" -o "$OUT")
[[ -n "$MODEL" ]] && ARGS+=(-m "$MODEL")

if [[ $RESUME -eq 1 ]]; then
  # resume は前回セッションのサンドボックス設定を引き継ぐ
  printf '%s' "$PROMPT" | codex exec resume --last - > "$OUT.log" 2>&1 || { tail -40 "$OUT.log" >&2; exit 1; }
  # resume では -o が使えない版があるため、ログ末尾を回答として扱う
  [[ -s "$OUT" ]] || tail -200 "$OUT.log" > "$OUT"
else
  printf '%s' "$PROMPT" | codex "${ARGS[@]}" - > "$OUT.log" 2>&1 || { tail -40 "$OUT.log" >&2; exit 1; }
fi

echo "=== Codex の回答 ($OUT) ==="
cat "$OUT"
