# インポート
import spacy
import docx
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
import re
import os
import pandas as pd
from bs4 import BeautifulSoup
import requests
import time
import warnings
import streamlit as st
# gmailを使ったメール送信用
from smtplib import SMTP
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

# 日本語モデルの構築
# GINZAの日本語辞書データ
nlp = spacy.load('ja_ginza_electra')

# テキストを処理
def read_text(file_name):
    if file_name[-4:] == 'docx' or file_name[-3:] == 'doc':
        doc = docx.Document(file_name)
        text_list = []
        for e in doc.element.body.iterchildren():
            if isinstance(e, CT_P):
                text_list.append(Paragraph(e, doc).text)
            if isinstance(e, CT_Tbl):
                for row in Table(e, doc).rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            text_list.append(p.text)
        text = ', '.join(text_list)
        return text        
    elif file_name[-3:] == 'txt':
        with open(file_name, mode='r', encoding='utf-8') as file:
            text = file.read()
            return text
    else:
        print('テキストを読み込めません')
        return

# 固有表現抽出結果の表示
def named_entity_recognition(text):
    if text:
        new_doc = nlp(text)
        entities = []
        for entity in new_doc.ents:
            # print(entity.text, entity.label_, entity.start_char, entity.end_char)
            entities.append([entity.text, entity.label_, entity.start_char, entity.end_char])
        return entities

# 固有名詞をデータフレーム化
def make_df(entities):
    df_name = pd.DataFrame(entities, columns=['text', 'label', 'start_char', 'end_char'])
    return df_name

# ヤフーニュースの検索バーのurlに変数keywordを追加する
def yahoo_news_search(keyword):
    url = f'https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8'
    # リンクを呼び出すコード
    r = requests.get(url)
    time.sleep(3)  # 負荷をかけないよう3秒待つ
    soup = BeautifulSoup(r.text, 'html.parser')
    count = soup.find_all('span')[1].string
    return count

# count数が0はピンク、10未満はオレンジに色付け
def highlight(df):
  styles = df.copy()  
  # DataFrame全体を初期化してから色付け
  styles.loc[:,:] = ''
  styles.loc[df['count'] < 10, :] = 'background-color: orange'
  styles.loc[df['count'] == 0, :] = 'background-color: pink'
  return styles

# データフレームに件数の列を追加
def count_df(df):
    df['count'] = df['text'].map(yahoo_news_search)
    # 件数を数値に変換
    df['count'] = df['count'].replace(',', '', regex=True)
    df['count'] = df['count'].astype('int')
    # applyの関数highlightでスタイル設定
    df_name_count = df.style.apply(highlight, axis=None)
    return df_name_count

# Excelファイルを保存
def df_to_excel(df):
    df.to_excel('caution_words.xlsx') 

gmail_address = st.secrets['gmail_address']
gmail_pass = st.secrets['gmail_pass']
    
# メールを送る
def sendGmailAttach(my_address, file_name, gmail_address, gmail_pass):
    if my_address == '':
        print('メールアドレスを入力してください')
    else:
        sender, password = gmail_address, gmail_pass # 送信元メールアドレスとgmailへのログイン情報 
        to = my_address  # 送信先メールアドレス
        sub = '注意ワードのチェック' #メール件名
        text = '原稿をテキスト分析して固有表現を抽出、Yahoo!ニュースとの一致度をチェックしました'  # メール本文
        html = """
        <html>
          <head></head>
          <body>
            <p>原稿をテキスト分析して、固有表現を抽出しました。</p>
            <p>Yahoo!ニュースのキーワード検索結果の件数(count)から一致度をチェック。<br>
            <p>カウント0(ピンク)、カウント10未満(オレンジ)は要確認ワードです。<br>
            <a href= {URL}>Yahoo!ニュース</a>
            </p> 
          </body>
        </html>
        """.format(URL='https://news.yahoo.co.jp/')

        host, port = 'smtp.gmail.com', 587

        # メールヘッダー
        msg = MIMEMultipart('alternative')
        msg['Subject'] = sub
        msg['From'] = sender
        msg['To'] = to

        # メール本文
        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')
        msg.attach(part1)
        msg.attach(part2)

        # 添付ファイル1の設定(Wordファイル)
        attach_file1 = {'name': f'{file_name}', 'path': f'{file_name}'} 
        attachment1 = MIMEBase('application', 'docx')
        file1 = open(attach_file1['path'], 'rb+')
        attachment1.set_payload(file1.read())
        file1.close()
        encoders.encode_base64(attachment1)
        attachment1.add_header("Content-Disposition", "attachment", filename=attach_file1['name'])
        msg.attach(attachment1)

        # 添付ファイル2の設定(Excelファイル)
        attach_file2 = {'name': 'caution_words.xlsx', 'path': 'caution_words.xlsx'} 
        attachment2 = MIMEBase('application', 'xlsx')
        file2 = open(attach_file2['path'], 'rb+')
        attachment2.set_payload(file2.read())
        file2.close()
        encoders.encode_base64(attachment2)
        attachment2.add_header("Content-Disposition", "attachment", filename=attach_file2['name'])
        msg.attach(attachment2)    

        # gmailへ接続(SMTPサーバーとして使用)
        gmail=SMTP("smtp.gmail.com", 587)
        gmail.starttls() # SMTP通信のコマンドを暗号化し、サーバーアクセスの認証を通す
        gmail.login(sender, password)
        gmail.send_message(msg)

# ファイルありメイン関数
def main_file(my_address, file_name, gmail_address, gmail_pass):
    text = read_text(file_name)
    entities =  named_entity_recognition(text)
    df_name = make_df(entities)
    df_name_count = count_df(df_name)
    df_to_excel(df_name_count)
    sendGmailAttach(my_address, file_name, gmail_address, gmail_pass)
    return df_name_count

# カレントディレクトリを取得
cwd_name = os.getcwd()

# テキスト入力時のメイン関数
def main_text(my_address, text, gmail_address, gmail_pass):
    entities =  named_entity_recognition(text)
    df_name = make_df(entities)
    df_name_count = count_df(df_name)
    df_to_excel(df_name_count)
    with open('./text_check.txt', mode='w', encoding='utf-8') as file:
        file.write(text)
    text_file = os.path.join(cwd_name, 'text_check.txt')
    sendGmailAttach(my_address, text_file, gmail_address, gmail_pass)
    return df_name_count

# streamlit画面作成
st.title("原稿の注意ワードチェック")
st.write("テキスト入力 または ファイル登録して、メールアドレスを入力してください")

# テキスト入力エリア
text = st.text_area("テキストを入力")

# ファイルアップロード
file = st.file_uploader("ファイル（Word[.doc, .docx] or テキスト[.txt]）をアップロード", accept_multiple_files= False)
if file:
    st.markdown(f'{file.name} をアップロードしました')
    file_name = file.name
    # ファイルを保存する
    with open(file_name, 'wb') as f:
        f.write(file.read())

# メールアドレス入力
my_address = st.text_input("メールアドレス")

# 実行ボタン作成
if st.button('処理を実行'):
    try:
        if file and my_address:
            st.write('処理を開始します')
            df = main_file(my_address, file_name, gmail_address, gmail_pass)
            st.dataframe(df)
            st.write('処理を終了しました')
        elif text and my_address:
            st.write('処理を開始します')
            df = main_text(my_address, text, gmail_address, gmail_pass)
            st.dataframe(df)
            st.write('処理を終了しました')
        else:
            st.write('入力内容を確認してください')
    except Exception as e:
        st.write('エラーが発生しました')
        st.write(e)
