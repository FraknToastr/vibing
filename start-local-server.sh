#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="/home/n0rt/Projects/vibing"
HOST="${HOST:-127.0.0.1}"
PORT="${1:-8000}"

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 is required but was not found in PATH." >&2
  exit 1
fi

if [ ! -d "$REPO_DIR" ]; then
  echo "Repo directory not found: $REPO_DIR" >&2
  exit 1
fi

if ! [[ "$PORT" =~ ^[0-9]+$ ]] || [ "$PORT" -lt 1 ] || [ "$PORT" -gt 65535 ]; then
  echo "Invalid port: $PORT (must be 1-65535)" >&2
  exit 1
fi

find_open_port() {
  local host="$1"
  local start_port="$2"

  python3 - "$host" "$start_port" <<'PY'
import socket
import sys

host = sys.argv[1]
start_port = int(sys.argv[2])

for port in range(start_port, 65536):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind((host, port))
        sock.close()
        print(port)
        sys.exit(0)
    except OSError:
        sock.close()
        continue

sys.exit(1)
PY
}

SELECTED_PORT="$(find_open_port "$HOST" "$PORT")"
if [ -z "$SELECTED_PORT" ]; then
  echo "No available port found on $HOST from $PORT to 65535." >&2
  exit 1
fi

if [ "$SELECTED_PORT" != "$PORT" ]; then
  echo "Port $PORT is already in use on $HOST. Using $SELECTED_PORT instead."
fi

cd "$REPO_DIR"
echo "Serving $REPO_DIR at http://$HOST:$SELECTED_PORT/"
exec python3 -m http.server "$SELECTED_PORT" --bind "$HOST"
