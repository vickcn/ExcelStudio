#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
詳細診斷本地LLM連線問題
"""

import requests
import socket
import os
from urllib.parse import urlparse
from dotenv import load_dotenv

load_dotenv()

def test_basic_connectivity():
    """測試基本網路連線"""
    llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1')
    
    print("=== 基本網路連線測試 ===")
    print(f"目標URL: {llm_api_url}")
    
    # 解析URL
    parsed = urlparse(llm_api_url)
    host = parsed.hostname
    port = parsed.port or (443 if parsed.scheme == 'https' else 80)
    
    print(f"主機: {host}")
    print(f"埠號: {port}")
    
    # 測試TCP連線
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(5)
        result = sock.connect_ex((host, port))
        sock.close()
        
        if result == 0:
            print("✓ TCP連線成功")
        else:
            print(f"✗ TCP連線失敗，錯誤碼: {result}")
            return False
    except Exception as e:
        print(f"✗ TCP連線測試失敗: {e}")
        return False
    
    return True

def test_http_methods():
    """測試不同的HTTP方法"""
    llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1')
    
    print("\n=== HTTP方法測試 ===")
    
    # 測試GET請求到根路徑
    try:
        base_url = llm_api_url.replace('/v1', '').replace('/chat/completions', '')
        print(f"測試GET請求到: {base_url}")
        
        response = requests.get(base_url, timeout=10)
        print(f"GET狀態碼: {response.status_code}")
        print(f"回應內容: {response.text[:200]}...")
        
    except Exception as e:
        print(f"GET請求失敗: {e}")
    
    # 測試OPTIONS請求
    try:
        api_url = llm_api_url
        if not api_url.endswith('/chat/completions'):
            if api_url.endswith('/v1'):
                api_url = f"{api_url}/chat/completions"
            else:
                api_url = f"{api_url}/v1/chat/completions"
        
        print(f"測試OPTIONS請求到: {api_url}")
        response = requests.options(api_url, timeout=10)
        print(f"OPTIONS狀態碼: {response.status_code}")
        print(f"允許的方法: {response.headers.get('Allow', 'N/A')}")
        
    except Exception as e:
        print(f"OPTIONS請求失敗: {e}")

def test_different_models():
    """測試不同的模型名稱"""
    llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1')
    original_model = os.getenv('LLM_MODEL_NAME', 'llama3.1')
    
    print("\n=== 模型測試 ===")
    
    # 處理API URL
    api_url = llm_api_url
    if not api_url.endswith('/chat/completions'):
        if api_url.endswith('/v1'):
            api_url = f"{api_url}/chat/completions"
        else:
            api_url = f"{api_url}/v1/chat/completions"
    
    # 測試不同的模型名稱
    test_models = [
        original_model,
        'llama3.1',
        'llama3',
        'gpt-3.5-turbo',  # 有些本地服務支援這個名稱
        'default'
    ]
    
    for model in test_models:
        print(f"\n測試模型: {model}")
        try:
            response = requests.post(
                api_url,
                json={
                    "model": model,
                    "messages": [{"role": "user", "content": "Hi"}],
                    "max_tokens": 10,
                    "temperature": 0.1
                },
                timeout=15,
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                }
            )
            
            print(f"  狀態碼: {response.status_code}")
            if response.status_code == 200:
                try:
                    result = response.json()
                    print(f"  成功! 回應: {result.get('choices', [{}])[0].get('message', {}).get('content', 'N/A')}")
                    return True
                except:
                    print(f"  JSON解析失敗: {response.text[:100]}")
            else:
                print(f"  失敗: {response.text[:200]}")
                
        except Exception as e:
            print(f"  請求失敗: {e}")
    
    return False

def test_alternative_endpoints():
    """測試替代的API端點"""
    llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1')
    
    print("\n=== 替代端點測試 ===")
    
    base_url = llm_api_url.replace('/v1', '').replace('/chat/completions', '')
    
    # 常見的API端點
    endpoints = [
        f"{base_url}/api/generate",  # Ollama風格
        f"{base_url}/api/chat",      # 另一種Ollama風格
        f"{base_url}/v1/completions", # OpenAI completions
        f"{base_url}/completions",    # 簡化版
        f"{base_url}/generate",       # 直接生成
    ]
    
    for endpoint in endpoints:
        print(f"\n測試端點: {endpoint}")
        try:
            # 測試簡單的POST請求
            response = requests.post(
                endpoint,
                json={
                    "prompt": "Hello",
                    "max_tokens": 10
                },
                timeout=10
            )
            
            print(f"  狀態碼: {response.status_code}")
            if response.status_code in [200, 201]:
                print(f"  成功! 回應: {response.text[:100]}")
            else:
                print(f"  回應: {response.text[:100]}")
                
        except Exception as e:
            print(f"  請求失敗: {e}")

def main():
    """主診斷函數"""
    print("開始詳細診斷本地LLM連線問題...\n")
    
    # 基本連線測試
    if not test_basic_connectivity():
        print("\n❌ 基本網路連線失敗，請檢查:")
        print("1. 服務是否正在運行")
        print("2. IP地址和埠號是否正確")
        print("3. 防火牆設定")
        return
    
    # HTTP方法測試
    test_http_methods()
    
    # 模型測試
    if test_different_models():
        print("\n✅ 找到可用的模型配置!")
        return
    
    # 替代端點測試
    test_alternative_endpoints()
    
    print("\n❌ 所有測試都失敗了，建議:")
    print("1. 檢查本地LLM服務是否正確啟動")
    print("2. 確認API格式是否符合OpenAI標準")
    print("3. 檢查服務日誌")
    print("4. 嘗試使用curl直接測試")

if __name__ == '__main__':
    main()
