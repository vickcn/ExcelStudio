#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
簡單的LLM測試腳本
"""

import os
import json
import requests
import openai
from dotenv import load_dotenv

load_dotenv()

def test_openai():
    """測試OpenAI API"""
    openai_api_key = os.getenv('OPENAI_API_KEY')
    
    if not openai_api_key or openai_api_key == 'your_openai_api_key_here':
        print("❌ OpenAI API金鑰未設定")
        return
    
    print(f"測試OpenAI API...")
    
    try:
        client = openai.OpenAI(api_key=openai_api_key)
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": "Hello, please respond with 'OpenAI API is working'"}
            ],
            max_tokens=50,
            temperature=0.7
        )
        
        content = response.choices[0].message.content
        print(f"✅ OpenAI回應: {content}")
        return True
        
    except Exception as e:
        print(f"❌ OpenAI API測試失敗: {e}")
        return False

def simple_llm_test():
    """簡單的LLM測試"""
    llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1').rstrip('/')
    llm_model_name = os.getenv('LLM_MODEL_NAME', 'llama3.1')
    api_key = os.getenv('LLM_API_KEY')
    
    print(f"測試LLM: {llm_api_url}")
    print(f"模型: {llm_model_name}")
    
    # 建立session
    session = requests.Session()
    base_headers = {
        'Content-Type': 'application/json',
        'User-Agent': 'LLM-API-Client/1.0'
    }
    session.headers.update(base_headers)
    
    # 非常簡單的測試訊息
    test_messages = [
        {"role": "user", "content": "Hello"}
    ]
    
    payload = {
        'model': llm_model_name,
        'messages': test_messages,
        'temperature': 0.7,
        'max_tokens': 50  # 很小的token數量
    }
    
    # 準備標頭
    headers = {}
    if api_key:
        headers['Authorization'] = f'Bearer {api_key}'
    
    print(f"發送請求...")
    print(f"Payload: {json.dumps(payload, indent=2, ensure_ascii=False)}")
    
    try:
        response = session.post(
            llm_api_url,
            json=payload,
            headers=headers,
            timeout=120  # 2分鐘超時
        )
        
        print(f"狀態碼: {response.status_code}")
        print(f"回應標頭: {dict(response.headers)}")
        print(f"回應長度: {len(response.text)}")
        print(f"回應內容: {response.text}")
        
        if response.status_code == 200:
            if response.text.strip():
                try:
                    result = response.json()
                    print(f"JSON解析成功: {json.dumps(result, indent=2, ensure_ascii=False)}")
                    
                    if 'choices' in result and len(result['choices']) > 0:
                        content = result['choices'][0]['message']['content']
                        print(f"✅ LLM回應: {content}")
                    elif 'response' in result:
                        print(f"✅ LLM回應: {result['response']}")
                    else:
                        print(f"❌ 未知回應格式")
                        
                except json.JSONDecodeError as e:
                    print(f"❌ JSON解析失敗: {e}")
            else:
                print(f"❌ 回應為空")
        else:
            print(f"❌ HTTP錯誤: {response.status_code}")
            
    except Exception as e:
        print(f"❌ 請求失敗: {e}")

if __name__ == "__main__":
    print("=== LLM測試 ===")
    
    # 測試OpenAI
    print("\n1. 測試OpenAI API:")
    openai_success = test_openai()
    
    # 測試本地LLM
    print("\n2. 測試本地LLM:")
    simple_llm_test()
    
    # 建議
    print("\n=== 建議 ===")
    if openai_success:
        print("✅ OpenAI API可用，建議使用OpenAI進行表格偵測")
        print("   執行: python table_detector.py correct_simple.xlsx --mode pure_numeric")
    else:
        print("❌ OpenAI API不可用，請設定OPENAI_API_KEY環境變數")
