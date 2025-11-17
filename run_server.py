import threading
import webbrowser
import time
from app import app

def run_flask():
    app.run(host="127.0.0.1", port=5000, debug=False)

if __name__ == "__main__":
    # 在后台线程启动 Flask
    flask_thread = threading.Thread(target=run_flask)
    flask_thread.daemon = True
    flask_thread.start()

    # 等 Flask 启动
    time.sleep(1.5)

    # 自动打开浏览器
    webbrowser.open("http://127.0.0.1:5000")

    # 阻止程序退出
    flask_thread.join()
