from selenium import webdriver
from time import sleep
# from urllib.request import urlopen as uReq

usr = input('Enter Email Id:')
#pwd = input('Enter password:')

zomatoDriver = webdriver.Chrome()
zomatoDriver.get('https://zomato.com/bangalore')
print('zomato opened')
sleep(5)

signing_box = zomatoDriver.find_element_by_xpath("//a[@id='signin-link']")
signing_box.click()
sleep(5)

window_before = zomatoDriver.window_handles[0]
print('window_before')
print('login popup appeared')

zomatoDriver.switch_to.frame(zomatoDriver.find_element_by_xpath(
    "//a[@id='login-email'] // span[@class='fontsize3']").click())

zomatoDriver.switch_to.default_content()
username_box = zomatoDriver.find_element_by_xpath(
    "//input[@type='text' and @class='zomato-form-input-plain' and @id='ld-email']")

username_box.clear()
username_box.send_keys(usr)

print("Email ID entered")
sleep(1)

next_thing = zomatoDriver.find_element_by_xpath(
    "//input[@type='submit' and @id='ld-submit-global']")
next_thing.click()

password_box = zomatoDriver.find_element_by_xpath(
    "//input[@placeholder='Enter One Time Password']")
sleep(3)
password_box.clear()

#usr1 = input('Enter your email:')
pwd1 = input('ENter the Password:')

outlookDriver = webdriver.Chrome()
outlookDriver.get('https://outlook.live.com/owa/')
print('outlook opened')

main_page = outlookDriver.current_window_handle

sleep(3)

sigining_box1 = outlookDriver.find_element_by_xpath(
    "//nav[@class='auxiliary-actions']//a[@data-task='signin']")
sigining_box1.click()
sleep(3)

outlookDriver.find_element_by_xpath("//input[@name='loginfmt']")

login_page = outlookDriver.current_window_handle
for handle in outlookDriver.window_handles:
    if handle != main_page:
        login_page = handle

outlookDriver.switch_to.window(login_page)

username_box1 = outlookDriver.find_element_by_xpath(
    "//input[@name='loginfmt']")
username_box1.clear()
username_box1.send_keys(usr)

print('email id entered')
sleep(2)

outlookDriver.find_element_by_xpath("//input[@type='submit']").click()

pwd1_box = outlookDriver.find_element_by_xpath("//input[@type='password']")
pwd1_box.clear()
pwd1_box.send_keys(pwd1)

print('password entered')
sleep(3)


final_click = outlookDriver.find_element_by_xpath("//input[@type='submit']")
final_click.click()
sleep(6)


outlookDriver.find_element_by_xpath(
    "(//span[contains(text(),'OTP')])[1]").click()
my_value = outlookDriver.find_element_by_xpath(
    "//p[@class='x_text-center x_zn-fontbig x_zn-bold']")
new_value = my_value.text
print(new_value)
#outlookDriver.quit()
print('Done!')


password_box.send_keys(new_value)
print("Password entered")

login_box = zomatoDriver.find_element_by_xpath(
    "//span[@class='verif-code-submit-text fontsize3' and text()='Go']")
login_box.click()

zomatoDriver.switch_to_window(window_before)
print("Done")
#input("Press any key to quit")
#zomatoDriver.quit()
print("Finished")
