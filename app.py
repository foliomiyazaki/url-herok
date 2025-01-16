import streamlit as st
import openai
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import pandas as pd
import time
from tqdm import tqdm

# 環境変数からOpenAI APIキーを取得
openai.api_key = os.getenv("OPENAI_API_KEY")

# GPT解析関数
def analyze_text_with_gpt(text):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Extract company information as structured data."},
                {"role": "user", "content": f"以下のテキストから情報を抽出:\n{text}"}
            ]
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        return f"GPTエラー: {e}"

# 会社情報ページ特定
def find_company_info_page(domain_url):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(domain_url)
    time.sleep(3)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()

    links = [urljoin(domain_url, a['href']) for a in soup.find_all('a', href=True)]
    return links[0] if links else domain_url

# Webページをスクレイピング
def scrape_page(url):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(url)
    time.sleep(3)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()
    return soup.get_text(separator="\n", strip=True)

# アップロードファイル処理
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    results = []

    st.write("Processing...")

    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="Processing URLs"):
        url = row.get('URL', '')
        company_info_url = find_company_info_page(url)
        page_text = scrape_page(company_info_url)
        gpt_result = analyze_text_with_gpt(page_text)

        results.append({
            "URL": url,
            "Company Info URL": company_info_url,
            "GPT Result": gpt_result
        })

    result_df = pd.DataFrame(results)
    st.write(result_df)

    # 結果をダウンロード可能に
    output_file = "results.xlsx"
    result_df.to_excel(output_file, index=False)
    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Results",
            data=file,
            file_name="results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
