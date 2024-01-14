# 必要なライブラリをインポート

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass


# ヘッドレスモードで起動するためのオプションを設定
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

# Chromeを立ち上げる
chrome_driver = webdriver.Chrome(options=chrome_options)

# TERAKOYAのトップページにアクセス
chrome_driver.get('https://terakoya.sejuku.net/register')

# 最大30秒間、ログインボタンが表示されるのを待つ
wait = WebDriverWait(chrome_driver, 30)
header_login_button = wait.until(
    EC.visibility_of_element_located(
        (By.CSS_SELECTOR,
         '#root > header > div.sc-jgyNOb.eGDJQH')
    )
)

# ログインボタンをクリックする
header_login_button.click()


# メールアドレスとパスワードを入力する

email_address = input('メールアドレスを入力してください: ')
password = getpass('パスワードを入力してください: ')


# メールアドレスとパスワードの入力欄を見つける

parent_element = chrome_driver.find_element(
    By.CSS_SELECTOR, '.sc-pyfCe.lduieC')
email_input = parent_element.find_element(By.NAME, 'email')
password_input = parent_element.find_element(By.NAME, 'password')


# メールアドレスとパスワードを設定

email_input.send_keys(email_address)
password_input.send_keys(password)


# ログインボタンをクリックしてログイン

form_login_button = wait.until(
    EC.visibility_of_element_located(
        (By.CSS_SELECTOR,
         '#root > div.sc-pyfCe.lduieC > div.sc-kDvujY.eCJBhf > '
         'div.sc-eDWCr.dePrRH > button')
    )
)

form_login_button.click()


# ログイン後、要素が読み込まれるまで待つ
wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'nav > img ')))

# スクリーンショットを撮る
chrome_driver.save_screenshot('screenshot.png')


chrome_driver.quit()
