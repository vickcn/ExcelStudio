#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM連接診斷工具
檢查地端LLM和OpenAI連接狀況
"""

import os
import sys
import requests
import json
from dotenv import load_dotenv
import openai

# 載入環境變數
load_dotenv()

def test_local_llm():
    """測試地端LLM連接"""
    print("=== 測試地端LLM連接 ===")
    
    llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1').rstrip('/')
    llm_model_name = os.getenv('LLM_MODEL_NAME', 'llama3.1')
    api_key = os.getenv('LLM_API_KEY')
    
    print(f"API URL: {llm_api_url}")
    print(f"模型名稱: {llm_model_name}")
    print(f"API金鑰: {'已設定' if api_key else '未設定'}")
    
    # 常見的地端LLM端點
    common_endpoints = [
        f"{llm_api_url}/chat/completions",
        f"{llm_api_url}/v1/chat/completions",
        "http://localhost:11434/api/chat",  # Ollama
        "http://localhost:11434/v1/chat/completions",  # Ollama OpenAI兼容
        "http://localhost:8080/v1/chat/completions",  # 其他常見端點
        "http://localhost:5000/v1/chat/completions",
    ]
    
    session = requests.Session()
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'User-Agent': 'MathRuleDetector/1.0'
    }
    
    if api_key:
        headers['Authorization'] = f'Bearer {api_key}'
    
    session.headers.update(headers)
    
    for endpoint in common_endpoints:
        print(f"\n測試端點: {endpoint}")
        try:
            # 先測試基本連接
            response = session.get(endpoint.replace('/chat/completions', '/models'), timeout=5)
            print(f"  模型列表請求: {response.status_code}")
            
            # 測試聊天完成
            test_data = {
                "model": llm_model_name,
                "messages": [
                    {"role": "user", "content": "Hello, test message"}
                ],
                "max_tokens": 10,
                "temperature": 0.1
            }
            
            response = session.post(endpoint, json=test_data, timeout=10)
            print(f"  聊天完成請求: {response.status_code}")
            
            if response.status_code == 200:
                print(f"  ✅ 成功！端點可用: {endpoint}")
                try:
                    result = response.json()
                    if 'choices' in result and len(result['choices']) > 0:
                        print(f"  回應內容: {result['choices'][0].get('message', {}).get('content', 'N/A')[:50]}...")
                        return endpoint
                except:
                    pass
            else:
                print(f"  ❌ 失敗: {response.text[:100]}")
                
        except requests.exceptions.ConnectionError:
            print(f"  ❌ 連接失敗: 無法連接到服務")
        except requests.exceptions.Timeout:
            print(f"  ❌ 超時: 服務響應太慢")
        except Exception as e:
            print(f"  ❌ 錯誤: {e}")
    
    return None

def test_openai():
    """測試OpenAI連接"""
    print("\n=== 測試OpenAI連接 ===")
    
    api_key = os.getenv('OPENAI_API_KEY')
    
    if not api_key or api_key == 'your_openai_api_key_here':
        print("❌ OpenAI API金鑰未設定")
        return False
    
    print(f"API金鑰: {api_key[:10]}...{api_key[-4:]}")
    
    try:
        client = openai.OpenAI(api_key=api_key)
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": "Hello, test message"}
            ],
            max_tokens=10,
            temperature=0.1
        )
        
        print("✅ OpenAI連接成功")
        print(f"回應: {response.choices[0].message.content}")
        return True
        
    except Exception as e:
        print(f"❌ OpenAI連接失敗: {e}")
        return False

def suggest_fixes():
    """建議修復方案"""
    print("\n=== 修復建議 ===")
    
    print("1. 地端LLM修復方案:")
    print("   - 確認Ollama或其他LLM服務正在運行")
    print("   - 檢查端口是否正確（通常是11434或8080）")
    print("   - 嘗試不同的API端點格式")
    print("   - 確認模型名稱正確")
    
    print("\n2. 環境變數設定:")
    print("   創建或更新 .env 文件:")
    print("   ```")
    print("   # Ollama設定")
    print("   LLM_API_URL=http://localhost:11434/v1")
    print("   LLM_MODEL_NAME=llama3.1")
    print("   # LLM_API_KEY=your_key_if_needed")
    print("   ")
    print("   # OpenAI設定")
    print("   OPENAI_API_KEY=your_openai_key")
    print("   ```")
    
    print("\n3. 快速解決方案:")
    print("   使用OpenAI API（如果已設定）:")
    print("   python run_math_rule_analysis.py --use-openai")

def create_fixed_detector():
    """創建修復版本的偵測器"""
    print("\n=== 創建修復版本 ===")
    
    # 測試可用的端點
    working_endpoint = test_local_llm()
    openai_works = test_openai()
    
    if working_endpoint:
        print(f"\n建議使用地端LLM端點: {working_endpoint}")
        
        # 更新環境變數建議
        base_url = working_endpoint.replace('/chat/completions', '')
        print(f"建議設定 LLM_API_URL={base_url}")
        
    elif openai_works:
        print("\n建議使用OpenAI API")
        print("執行: python run_math_rule_analysis.py --use-openai")
    else:
        print("\n❌ 沒有可用的LLM服務")
        print("請先設定地端LLM或OpenAI API")

def main():
    """主程式"""
    print("LLM連接診斷工具")
    print("=" * 50)
    
    # 顯示當前環境變數
    print("當前環境變數:")
    print(f"  LLM_API_URL: {os.getenv('LLM_API_URL', '未設定')}")
    print(f"  LLM_MODEL_NAME: {os.getenv('LLM_MODEL_NAME', '未設定')}")
    print(f"  LLM_API_KEY: {'已設定' if os.getenv('LLM_API_KEY') else '未設定'}")
    print(f"  OPENAI_API_KEY: {'已設定' if os.getenv('OPENAI_API_KEY') else '未設定'}")
    
    # 執行測試
    create_fixed_detector()
    
    # 提供建議
    suggest_fixes()

if __name__ == "__main__":
    main()
