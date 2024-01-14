# 必要なライブラリをインポート
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass


# chromeブラウザ起動オプションを指定
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

# Chromeを立ち上げる
chrome_driver = webdriver.Chrome(options=chrome_options)


# WEB申請システムのトップページにアクセス
chrome_driver.get('https://xs016594.xsrv.jp/public/login')


# メールアドレスとパスワードの入力欄を見つける
parent_element = chrome_driver.find_element(
    #    CSSセレクタを記述
    By.CSS_SELECTOR, '[name="email"]')
email_input = parent_element.find_element(By.NAME, 'email')
password_input = parent_element.find_element(By.NAME, 'password')

# メールアドレスとパスワードを入力する
email_address = 'test@example.com'
password = 'test01'

# メールアドレスとパスワードを設定
email_input.send_keys(email_address)
password_input.send_keys(password)


# 最大30秒間待つ
wait = WebDriverWait(chrome_driver, 30)

# ログインボタンをクリックしてログイン
form_login_button = wait.until(
    EC.visibility_of_element_located(
        (By.CSS_SELECTOR,
         #  ここを変更する
         '.mt-3.btn.webapplication-submit-button.w-100'
         )
    )
)

form_login_button.click()
