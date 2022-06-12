import pandas as pd
import json
import time

# auto chromedriver version finder

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.common.exceptions import NoSuchElementException

from termcolor import colored

from datetime import datetime

ROOT_URL, Username, Password = "", "", ""
driver = ""

datafile = ""
sheet_list = ""


def configure_variables():
    global ROOT_URL, Username, Password, datafile, sheet_list
    ROOT_URL = "https://sampoorna.kite.kerala.gov.in:446"
    Username = "admin@36095"
    Password = "gcs@sampoorna2018"

    datafile = "data/data.xlsx"
    sheet_list = "Sheet2"


def start_driver():
    global driver
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--disable-infobars")
    prefs = {"profile.default_content_setting_values.notifications": 1}
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


def login_gcs(WebUrl, user_name, pass_name):
    global driver
    print("Logging in...")

    driver.get(WebUrl)

    username = driver.find_element(By.XPATH, '//input[@id="user_username"]')
    username.click()
    username.send_keys(user_name)

    password = driver.find_element(By.XPATH, '//input[@id="user_password"]')
    password.click()
    password.send_keys(pass_name)

    login = driver.find_element(By.XPATH, '//input[@type="submit"]').click()

    # time.sleep(3)
    # opening admission
    driver.get(WebUrl + "/13572/student/admission")


def check_error():
    global driver

    try:
        check = driver.find_element(By.XPATH, '//div[@id="errorExplanation"]')
        print(check.get_attribute('innerHTML'))
        return [True, check.get_attribute('innerHTML')]
    except NoSuchElementException:
        return [False, ""]


def start_admitting(student_data):
    global driver, ROOT_URL
    errors = []
    issues = {}
    # temp = [3130, 2952, 3274, 3131, 2699, 2307, 2497, 2762, 3704]
    # temp_ind = [71, 74, 82, 105, 113, 126, 146, 167, 169]
    for i in range(162, len(student_data['Admn.No.'])):
    # for i in temp_ind:
        print("Starting ", student_data['Admn.No.'][i], " index = ", i)
        # Admission Number
        admnno = driver.find_element(By.XPATH, '//input[@id="student_admission_no"]')
        admnno.click()
        admnno.send_keys(student_data['Admn.No.'][i])

        # Student Name
        name = driver.find_element(By.XPATH, '//input[@id="student_full_name"]')
        name.click()
        name.send_keys(student_data['Student Full Name'][i])

        # gender
        if student_data['Gender'][i] == "Girl":
            driver.find_element(By.XPATH, '//input[@id="student_gender_f"]').click()
        elif student_data['Gender'][i] == "Boy":
            driver.find_element(By.XPATH, '//input[@id="student_gender_m"]').click()

        # Student UID
        try:
            uid = driver.find_element(By.XPATH, '//input[@id="student_uid"]')
            uid.click()
            uid.send_keys(int(student_data['Student UID (Aadhar No)'][i]))
        except ValueError:
            errors.append(student_data['Admn.No.'][i])

        # Mother name
        mname = driver.find_element(By.XPATH, '//input[@id="student_mother_full_name"]')
        mname.click()
        mname.send_keys(student_data['Name of Mother'][i])

        # Father name
        fname = driver.find_element(By.XPATH, '//input[@id="student_father_full_name"]')
        fname.click()
        fname.send_keys(student_data['Name of Father'][i])

        # Annual Income
        anninc = driver.find_element(By.XPATH, '//input[@id="student_annual_income"]')
        anninc.click()
        anninc.send_keys(int(student_data['Annual Income'][i]))

        # BPL/APL
        if student_data['BPL/APL'][i] == "APL":
            driver.find_element(By.XPATH, '//input[@id="student_is_apl_true"]').click()
        elif student_data['BPL/APL'][i] == "BPL":
            driver.find_element(By.XPATH, '//input[@id="student_is_apl_false"]').click()

        ## Address details
        # House Name
        hn = student_data['House Name'][i].replace(",", "")
        hn = hn.replace("  ", "")
        temp = driver.find_element(By.XPATH, '//input[@id="student_address_line1"]')
        temp.click()
        temp.send_keys(hn)

        # Street/Place
        sp = student_data['Street/Place'][i].replace(",", "")
        sp = sp.replace("  ", " ")
        temp = driver.find_element(By.XPATH, '//input[@id="student_address_line2"]')
        temp.click()
        temp.send_keys(sp)

        # Post Office
        temp = driver.find_element(By.XPATH, '//input[@id="student_postoffice"]')
        temp.click()
        temp.send_keys(student_data['Post Office'][i])

        # Pin Code
        temp = driver.find_element(By.XPATH, '//input[@id="student_pincode"]')
        temp.click()
        temp.send_keys(student_data['Pin Code'][i])

        # Mobile
        temp = driver.find_element(By.XPATH, '//input[@id="student_phone1"]')
        temp.click()
        temp.send_keys(int(student_data['Mobile Number'][i]))

        # Email
        temp = driver.find_element(By.XPATH, '//input[@id="student_email"]')
        temp.click()
        temp.send_keys(student_data['Email address'][i])

        # DOA
        ddmmyyyy = str(student_data['Date of Admn.'][i])

        if ":" in ddmmyyyy:
            ddmmyyyy = ddmmyyyy.replace(" 00:00:00", "")
            ddmmyyyy = ddmmyyyy.split("-")
            ddmmyyyy.reverse()
        else:
            ddmmyyyy = ddmmyyyy.split("-")

        print(ddmmyyyy)

        # dd
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_admission_date_3i"]'))
        select.select_by_value(str(int(ddmmyyyy[0])))

        # mm
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_admission_date_2i"]'))
        select.select_by_value(str(int(ddmmyyyy[1])))

        # yyyy
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_admission_date_1i"]'))
        select.select_by_value(str(int(ddmmyyyy[2])))

        # DOB
        ddmmyyyy = str(student_data['Date of Birth'][i])
        if ":" in ddmmyyyy:
            ddmmyyyy = ddmmyyyy.replace(" 00:00:00", "")
            ddmmyyyy = ddmmyyyy.split("-")
            ddmmyyyy.reverse()
        else:
            ddmmyyyy = ddmmyyyy.split("-")

        print(ddmmyyyy)

        # dd
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_date_of_birth_3i"]'))
        select.select_by_value(str(int(ddmmyyyy[0])))

        # mm
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_date_of_birth_2i"]'))
        select.select_by_value(str(int(ddmmyyyy[1])))

        # yyyy
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_date_of_birth_1i"]'))
        select.select_by_value(str(int(ddmmyyyy[2])))

        # Place of Birth
        pob = str(student_data['Place of Birth'][i]).replace(",", "")
        pob = pob.replace(".", "")
        pob = pob.replace("&", "")
        pob = pob.replace("/", " ")
        pob = pob.replace("  ", " ")
        temp = driver.find_element(By.XPATH, '//input[@id="student_birth_place"]')
        temp.click()
        temp.send_keys(pob)

        # Blood Group
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_blood_group"]'))
        select.select_by_value(student_data['Blood Group'][i])

        # Religion
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_religion_id"]'))
        select.select_by_visible_text(student_data['Religion'][i])

        # Category
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_reservation_group_id"]'))
        select.select_by_visible_text(student_data['Category'][i])

        # Caste
        temp = driver.find_element(By.XPATH, '//input[@id="student_caste_name"]')
        temp.click()
        temp.send_keys(student_data['Caste'][i])

        # std on admission
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_class_on_admission_id"]'))
        select.select_by_value(str(student_data['Class'][i]))

        # class
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_course_id"]'))
        select.select_by_value(str(student_data['Class'][i]))

        time.sleep(2)

        # Division
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_batch_id"]'))
        select.select_by_visible_text("G1 2021-2022")

        # First Lang
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_first_language_id"]'))
        select.select_by_value("2")

        # Additional Lang
        select = Select(driver.find_element(By.XPATH, '//select[@id="student_additional_language_id"]'))
        select.select_by_visible_text("Not Applicable")

        # Physical Challenge
        # select = Select(driver.find_element(By.XPATH, '//select[@id="student_physically_challenged"]'))
        # if student_data['Physical Challenge'][i] == "Yes":
        #     select.select_by_value('YES')

        # Vaccination done - unchecked
        driver.find_element(By.XPATH, '//input[@id="student_is_vaccinated"]').click()

        # todo : click submit
        # Check for errors
        try:
            s = driver.find_element(By.XPATH, '//span[@class=" LV_validation_message LV_invalid"]')
            print("Error !!!!", s.text)

            if s.text == "UID must be exactly 12 digits long":
                errors.append(student_data['Admn.No.'][i])
            else:
                n = input()
        except NoSuchElementException:
            pass


        n = input("Enter confirmation")
        if n == "y":
            save = driver.find_element(By.XPATH, '//input[@type="submit"]').click()
        else:
            pass



        # print("STOPPED RIGHT THERE !!")
        # n = input()
        while True:
            try:
                check = driver.find_element(By.XPATH, '//p[@class="flash-msg"]')
                print(check.text)
                if "Student saved" in check.text:
                    break
                else:
                    if not check_error()[0]:
                        continue
                    else:
                        issues[student_data['Admn.No.'][i]] = check_error()[1]
            except:
                if not check_error()[0]:
                    continue
                else:
                    issues[student_data['Admn.No.'][i]] = check_error()[1]
                    driver.get(ROOT_URL + "/13572/student/admission")
                    break

        time.sleep(3)

        if errors:
            save_error("error.txt", errors)
            errors = []
        if issues:
            save_error("issue.txt", issues)
            issues = {}

        # driver.refresh()
        #
        # break

    print(errors)
    print(issues)

    # save_error("error.txt", errors)
    # save_error("issue.txt", issues)

    # break


def get_data_from_table():
    global datafile, sheet_list
    df = pd.read_excel(datafile, sheet_name=sheet_list, engine='openpyxl')
    student_data = {}
    for i in enumerate(df.head(0)):
        student_data[i[1]] = (df[i[1]].values.tolist())
    print(colored("Data loaded.", 'green', attrs=['reverse', 'blink']))
    return student_data


def start_scrape():
    global sheet_list, sheet
    global ROOT_URL, Username, Password
    print("Starting Driver...")
    start_driver()
    print(colored("Driver Loaded Success.", 'green', attrs=['reverse', 'blink']))
    login_gcs(ROOT_URL, Username, Password)
    print(colored("Starting Admission.", 'green', attrs=['reverse', 'blink']))
    student_data = get_data_from_table()
    start_admitting(student_data)


def save_error(file, data):
    fo = open(file, "a")
    fo.write(str(data))
    fo.close()


configure_variables()
start_scrape()
