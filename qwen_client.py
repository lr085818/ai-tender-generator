import os
import requests
import logging

DEEPSEEK_API_KEY = os.getenv('DEEPSEEK_API_KEY', '')
DEEPSEEK_BASE_URL = os.getenv('DEEPSEEK_BASE_URL', 'https://api.deepseek.com')
DEEPSEEK_CHAT_MODEL = os.getenv('DEEPSEEK_CHAT_MODEL', 'deepseek-v4-flash')


def call_llm_api(messages, model=None, timeout=120):
    """调用 DeepSeek Chat API"""
    api_key = DEEPSEEK_API_KEY
    if not api_key:
        raise Exception("DEEPSEEK_API_KEY is not set")

    url = f'{DEEPSEEK_BASE_URL}/chat/completions'
    headers = {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    data = {
        'model': model or DEEPSEEK_CHAT_MODEL,
        'messages': messages,
    }
    logging.info(f"[LLM] DeepSeek {data['model']}，消息长度: {sum(len(m.get('content','')) for m in messages)} 字符")
    response = requests.post(url, headers=headers, json=data, timeout=timeout)
    if response.status_code != 200:
        error_message = f"DeepSeek API Error: Status Code: {response.status_code}, Response Body: {response.text}"
        logging.error(error_message)
        print(error_message)
    response.raise_for_status()
    return response.json()


call_dashscope_api = call_llm_api
