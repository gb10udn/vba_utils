import http.server
import socketserver
import json
from dataclasses import dataclass, asdict


@dataclass(frozen=True)
class CellInfo:
    """
    エクセル VBA 側に値を返す際のフォーマット。
    """
    x: int
    y: int
    value: str
    sheet_name: str = ''  # INFO: 240721 指定無しの場合、空文字である (!= None) 点に注意せよ。(None が VBA に送りにくい + エクセル空文字のシート名は作れない ため。)


class ServerHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self) -> None:
        if self.path == '/':
            # [START] json でデータ受信
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data)
            print(f'Received data: {data}')
            # [END] json でデータ受信

            # [START] メイン処理
            pass  # INFO: 240721 ビジネスロジック
            result = [
                CellInfo(x=1, y=1, value='this is test ...'),
                CellInfo(x=2, y=1, value='this is test !!!'),
                CellInfo(x=3, y=1, value='日本語いけるかな？'),
                CellInfo(x=4, y=1, value='改行\nコードは\nどうや？'),
            ]
            # [END] メイン処理

            # [START] 値を送出する
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps([asdict(res) for res in result], ensure_ascii=False).encode('utf-8'))  # TODO: 240720 エクセル座標、シート名、書き込む値を返す。(vba はそれを受け取り、値を書き込む責務を持つ。)
            # [END] 値を送出する
        
        elif self.path == '/kill':  # INFO: 240721 ここで、URL を分岐させれる。(kill プロセスとかを実行させるといいのかも？)
            pass

        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'{"error": "Not Found"}')
            return
    

if __name__ == '__main__':
    PORT = 8000  # TODO: 240721 呼び出し側が開いたポート確認して指定するのが良さそう。
    with socketserver.TCPServer(("127.0.0.1", PORT), ServerHandler) as httpd:
        print("serving at port", PORT)
        httpd.serve_forever()