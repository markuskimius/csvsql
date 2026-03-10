"""CLI entry point — serves the CSVSQL web app and opens a browser."""

import argparse
import os
import signal
import sys
import threading
import webbrowser
from functools import partial
from http.server import HTTPServer, SimpleHTTPRequestHandler


def main():
    parser = argparse.ArgumentParser(
        prog="csvsql",
        description="Browser-based CSV database with SQL query support.",
    )
    parser.add_argument(
        "-p", "--port", type=int, default=8000, help="port to serve on (default: 8000)"
    )
    parser.add_argument(
        "--no-browser", action="store_true", help="don't open a browser automatically"
    )
    parser.add_argument(
        "--host", default="127.0.0.1", help="host to bind to (default: 127.0.0.1)"
    )
    args = parser.parse_args()

    static_dir = os.path.join(os.path.dirname(__file__), "static")
    handler = partial(SimpleHTTPRequestHandler, directory=static_dir)

    # Try the requested port, increment if unavailable
    port = args.port
    for attempt in range(10):
        try:
            server = HTTPServer((args.host, port), handler)
            break
        except OSError:
            port += 1
    else:
        print(f"Error: Could not find an available port.", file=sys.stderr)
        sys.exit(1)

    url = f"http://{args.host}:{port}"
    print(f"Serving CSVSQL at {url}")
    print("Press Ctrl+C to stop.")

    if not args.no_browser:
        threading.Timer(0.5, webbrowser.open, args=(url,)).start()

    signal.signal(signal.SIGINT, lambda *_: sys.exit(0))

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
