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

# # Googleのトップページを開く
# chrome_driver, get('https://www.google.com')


# WEB申請システムのトップページにアクセス
chrome_driver.get('https://xs016594.xsrv.jp/public/login')

# 最大30秒間、ログインボタンが表示されるのを待つ
# wait = WebDriverWait(chrome_driver, 30)
# login_button = wait.until(
#     EC.visibility_of_element_located(
#         (By.CSS_SELECTOR,
#          #  ここを変更する
#          '#submit-button')
#     )
# )

# # ログインボタンをクリックする
# # login_button.click()


# メールアドレスとパスワードの入力欄を見つける
parent_element = chrome_driver.find_element(
    # ここを変更する
    By.CSS_SELECTOR, '.webapplication-login-input')
email_input = parent_element.find_element(By.NAME, 'email')
password_input = parent_element.find_element(By.NAME, 'password')

# メールアドレスとパスワードを入力する
email_address = input('メールアドレスを入力してください: ')
password = getpass('パスワードを入力してください: ')

# メールアドレスとパスワードを設定
email_input.send_keys(email_address)
password_input.send_keys(password)


# ログインボタンをクリックしてログイン
form_login_button = wait.until(
    EC.visibility_of_element_located(
        (By.CSS_SELECTOR,
         #  ここを変更する
         '.webapplication-submit-button'
         )
    )
)

form_login_button.click()


# ログイン後、要素が読み込まれるまで待つ
# wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ' ')))

# スクリーンショットを撮る
# chrome_driver.save_screenshot('screenshot.png')

# # 閉じる
# chrome_driver.quit()
