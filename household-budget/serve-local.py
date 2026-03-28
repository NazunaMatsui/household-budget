#!/usr/bin/env python3
"""
家計簿を http://localhost:5173 で配信する（Node 不要）。
終了: Ctrl+C
"""
from __future__ import annotations

import http.server
import mimetypes
import socketserver
import sys
from pathlib import Path

PORT = 5173
HOST = "127.0.0.1"

# PWA 用に .webmanifest の MIME を明示（未登録環境向け）
mimetypes.add_type("application/manifest+json", ".webmanifest", strict=False)


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(Path(__file__).resolve().parent), **kwargs)

    def log_message(self, format: str, *args) -> None:
        sys.stderr.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), format % args))


def main() -> None:
    with socketserver.TCPServer((HOST, PORT), Handler) as httpd:
        print(f"家計簿: http://localhost:{PORT}/  （終了は Ctrl+C）")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n停止しました。")


if __name__ == "__main__":
    main()
