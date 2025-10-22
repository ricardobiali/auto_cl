from http.server import HTTPServer, SimpleHTTPRequestHandler
import os
import user_data

class MyHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        # Modifica apenas o HTML principal
        if self.path in ('/', '/index.html'):
            with open("index.html", "r", encoding="utf-8") as f:
                html_template = f.read()
            html_content = html_template.replace(
                '<p class="mb-0" style="text-align: left; font-weight: bold;">Seja bem-vindo</p>',
                f'<p class="mb-0" style="text-align: left; font-weight: bold;">Seja bem-vindo, {user_data.full_name}!</p>'
            )
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
        else:
            # Para outros arquivos (CSS, JS, imagens), serve normalmente
            super().do_GET()

os.chdir(os.path.dirname(__file__))  # garante que caminhos relativos funcionem
port = 8000
server = HTTPServer(("localhost", port), MyHandler)
print(f"Servidor rodando em http://localhost:{port}")
server.serve_forever()