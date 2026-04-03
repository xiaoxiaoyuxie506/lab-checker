"""
启动 Flask 服务并创建 ngrok 隧道
"""
import os
import sys
import subprocess
import time
import threading

# 添加 pyngrok 路径
sys.path.insert(0, r'D:\Users\xiaol\AppData\Local\Programs\Python\Python312\Lib\site-packages')

def start_flask():
    """启动 Flask 服务"""
    os.chdir(r'C:\Users\xiaol\.stepclaw\workspace\lab-checker')
    subprocess.run([sys.executable, 'app.py'])

def start_ngrok():
    """启动 ngrok 隧道"""
    try:
        from pyngrok import ngrok
        
        # 等待 Flask 启动
        time.sleep(3)
        
        # 创建隧道
        public_url = ngrok.connect(5000, "http")
        print(f"\n" + "="*60)
        print(f"🎉 内网穿透成功！")
        print(f"📱 手机流量访问地址: {public_url}")
        print(f"🔗 本地服务地址: http://localhost:5000")
        print(f"="*60 + "\n")
        print(f"⚠️  注意: 这个地址每次重启都会变化")
        print(f"💡 按 Ctrl+C 停止服务\n")
        
        # 保持运行
        while True:
            time.sleep(1)
    except Exception as e:
        print(f"ngrok 启动失败: {e}")
        print("请检查网络连接，或尝试其他方案")

if __name__ == '__main__':
    # 启动 Flask（在后台线程）
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()
    
    # 启动 ngrok
    start_ngrok()
