import http.server
import socketserver
import json
import time
import socket


class ServerHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        data = json.loads(post_data)

        if self.path == '/':
            a = data['a']
            b = data['b']
            result = {'result': a + b}
        
        elif self.path == '/kill':  # INFO: 240721 ここで、URL を分岐させれる。(kill プロセスとかを実行させるといいのかも？)
            pass

        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'{"error": "Not Found"}')
            return

        
        # レスポンスの設定
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        
        self.wfile.write(json.dumps(result).encode('utf-8'))  # TODO: 240720 これで、同期的に結果を受け取れるみたい、、、、！


if __name__ == '__main__':
    PORT = 8001  # TODO: 240721 呼び出し側が開いたポート確認して指定するのが良さそう。
    with socketserver.TCPServer(("127.0.0.1", PORT), ServerHandler) as httpd:
        print("serving at port", PORT)
        httpd.serve_forever()