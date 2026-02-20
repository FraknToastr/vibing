#!/usr/bin/env bash
set -euo pipefail

PORT=${1:-8001}
STAMP=$(date -u +%Y%m%dT%H%M%SZ)
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

find_free_port() {
  python3 - "$1" <<'PY'
import socket
import sys

start = int(sys.argv[1])
for port in range(start, start + 100):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind(("127.0.0.1", port))
        except OSError:
            continue
        print(port)
        sys.exit(0)
print(start)
PY
}

PORT="$(find_free_port "${PORT}")"

cat <<MSG
Starting static server on port ${PORT}...
Serving from: ${SCRIPT_DIR}
Open (Windows): http://wsl.localhost:${PORT}/index.html?cb=${STAMP}
Open (WSL): http://localhost:${PORT}/index.html?cb=${STAMP}
MSG

python3 -m http.server "${PORT}" --directory "${SCRIPT_DIR}"
