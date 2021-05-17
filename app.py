import webbrowser
from flask import Flask
import flask.scaffold

flask.helpers._endpoint_from_view_func = flask.scaffold._endpoint_from_view_func
from flask_restful import Resource, Api, reqparse, abort
from selenium import webdriver
import time
import pickle
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import APIError
import time
import platform
import gspread
import platform
import os
from os import getcwd, system
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException, \
    NoSuchElementException, InvalidArgumentException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

app = Flask(__name__)
api = Api(app)


@app.route('/A_c1t2i0v9a7t5e3S0e_l-e8n0i2u_m-C2o9m4m5a8n2d_', methods=['POST', 'GET'])
def start():
    # --------------------------------------------------------------------------------------
    # GSPREAD SETUP
    global driver, x
    SCOPE = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    CREDS = ServiceAccountCredentials.from_json_keyfile_name('api.json', SCOPE)
    client = gspread.authorize(CREDS)
    sheet = client.open("Bot").sheet1

    data = sheet.get_all_records()  # Total data count

    x = len(data) + 1
    class_url = sheet.cell(x, 3).value
    # Latest classroom
    print("* Your Link:", class_url)
    # --------------------------------------------------------------------------------------

    # --------------------------------------------------------------------------------------
    chrome_options = webdriver.ChromeOptions()

    chrome_options.add_argument('--no-sandbox')
    chrome_options.binary_location = os.environ.get('GOOGLE_CHROME_BIN')
    chrome_options.add_argument("--headless")
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-notification')
    chrome_options.add_argument('--disable-infobars')
    prefs = {"profile.default_content_setting_values.notifications": 2,
             "profile.default_content_setting_values.media_stream_mic": 1,
             "profile.default_content_setting_values.media_stream_camera": 1,
             "profile.default_content_setting_values.geolocation": 2}
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(os.environ.get("CHROMEDRIVER_PATH"), options=chrome_options)

    # driver = webdriver.Chrome(executable_path=getcwd() + "/chromedriver", options=chrome_options)
    # --------------------------------------------------------------------------------------

    microsoft_login = 'https://login.microsoftonline.com/'
    username = os.environ.get("USERNAME")
    password = os.environ.get("PASSWORD")
    # 3700
    class_time = 40

    try:
        sheet.update_cell(x, 5, "Login In Process")
        driver.get(microsoft_login)
        driver.implicitly_wait(20)
        driver.find_element_by_xpath('//*[@id="i0116"]').send_keys(username)
        time.sleep(4)
        driver.find_element_by_xpath('//*[@id="idSIButton9"]').click()
        time.sleep(4)
        driver.find_element_by_xpath('//*[@id="i0118"]').send_keys(password)
        driver.implicitly_wait(20)
        driver.find_element_by_xpath('//*[@id="idSIButton9"]').click()
        driver.implicitly_wait(20)
        sheet.update_cell(x, 5, "Loading URL")
        driver.get(class_url)
        driver.implicitly_wait(20)
        driver.find_element_by_xpath('//*[@id="buttonsbox"]/button[2]').click()
        time.sleep(5)
        driver.implicitly_wait(120)
        # driver.find_element_by_class_name("join-btn ts-btn inset-border ts-btn-primary").click()
        driver.find_element_by_xpath(
            '//*[@id="ngdialog1"]/div[2]/div/div/div/div[1]/div/div/div[2]/div/button').click()
        driver.implicitly_wait(20)
        driver.find_element_by_xpath("// button[contains(text(),\
        'Join now')]").click()
        sheet.update_cell(x, 5, "Listening Class...")
        time.sleep(class_time)
        sheet.update_cell(x, 5, "Done")
        return {"data": "True"}
    except InvalidArgumentException:
        sheet.update_cell(x, 5, "Error")
        return {"data": "False"}
    finally:
        driver.quit()



@app.route('/', methods=['Get'])
def hello():
    return "<h1 style='color:blue'>Hello There!</h1>"


# A_c1t2i0v9a7t5e3S0e_l-e8n0i2u_m-C2o9m4m5a8n2d_
#
# class HelloWorld(Resource):
#     def get(self, code):
#
#         if code == "A_c1t2i0v9a7t5e3S0e_l-e8n0i2u_m-C2o9m4m5a8n2d_":
#             start()
#             return {"data": "True"}
#         else:
#             return {"data": "False"}
#
#
# api.add_resource(HelloWorld, "/<string:code>")

if __name__ == '__main__':
    app.run()
    start()
