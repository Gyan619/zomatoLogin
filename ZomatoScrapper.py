from selenium import webdriver
from time import sleep
from openpyxl import Workbook
from openpyxl.styles import Font, Color, colors
from openpyxl.utils import FORMULAE

usr = input('Enter email id:')
# pwd = input('Enter password:')

zomatoDriver = webdriver.Safari()
zomatoDriver.get('https://zomato.com/bangalore')
print('zomato opened')
sleep(5)
zomatoDriver.maximize_window()

signing_box = zomatoDriver.find_element_by_xpath(
    "//a[contains(text(),' in')]")
signing_box.click()
sleep(5)

window_before = zomatoDriver.window_handles[0]
print('window_before')
print('login popup appeared')

zomatoDriver.switch_to.frame(zomatoDriver.find_element_by_xpath(
    "//span[contains(text(),'Continue with Email')]").click())

zomatoDriver.switch_to.default_content()
username_box = zomatoDriver.find_element_by_xpath(
    "//input[@type='text' and @autocomplete='on']")

# username_box.clear()
username_box.send_keys(usr)

print(usr)
sleep(1)

next_thing = zomatoDriver.find_element_by_xpath(
    "//span[contains(text(),'Send OTP')]")
next_thing.click()

password_box = zomatoDriver.find_element_by_xpath(
    "//input[@type='text' and @autocomplete='on']")
sleep(3)
# password_box.clear()

# usr1 = input('Enter your email:')
# pwd1 = input('ENter the Password:')
pwd1 = input('Enter your password:')

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
    "(//span[@data-automationid='splitbuttonprimary'])[10]").click()
print('clicked on Other tab')
sleep(10)
outlookDriver.find_element_by_xpath(
    "(//span[contains(text(),'Zomato')])[1]").click()
my_value = outlookDriver.find_element_by_xpath(
    "//p[@class='x_text-center x_zn-fontbig x_zn-bold']")
new_value = my_value.text
print(new_value)
# outlookDriver.quit()
print('Done!')

outlookDriver.find_element_by_xpath("//img[@alt='GK']").click()
#outlookDriver.find_element_by_xpath("//a[contains(text(),'Sign out')]").click()

outlookDriver.close()

zomatoDriver.switch_to.window_before
print("Done")

sleep(7)

zomatoDriver.switch_to.default_content()
password_box.send_keys(new_value)
print("Password entered")

login_box = zomatoDriver.find_element_by_xpath(
    "//span[contains(text(),'Proceed')]")
login_box.click()


# input("Press any key to quit")
# zomatoDriver.quit()
# print("Finished")
# sleep(7)
dashboard = zomatoDriver.find_element_by_xpath(
    "//span[@class='username right ml0']")
dashboard.click()
sleep(3)

userProfile = zomatoDriver.find_element_by_xpath(
    "//a[contains(text(),'Profile')]")
userProfile.click()
sleep(4)

orderHistory = zomatoDriver.find_element_by_xpath(
    "//a[contains(text(),'Order History')]")
orderHistory.click()
print('My order history appears')
sleep(2)

currentUrl = zomatoDriver.current_url
print(currentUrl)

allOrder = zomatoDriver.find_element_by_xpath("//span[@class = 'grey-text']")
totalOrderNo = allOrder.text
print(totalOrderNo)
print(type(totalOrderNo))

for myOrder in range(int(totalOrderNo) // 20):
    zomatoDriver.find_element_by_xpath(
        "//button[@class = 'ui basic button']").click()
    sleep(8)
    zomatoDriver.execute_script(
        "window.scrollTo(0, document.body.scrollHeight);")


allRestaurantName = zomatoDriver.find_elements_by_xpath(
    "//div[@class='content']//div[@class='header nowrap']//a[@class='left mr10 ']")
# print(allRestaurantName)
RestaurantName = []
for restaurantName in allRestaurantName:
    RestaurantNames = restaurantName.text.strip()
    print(RestaurantNames)
    RestaurantName.append(RestaurantNames)

    # print(type(RestaurantName))

allAddress = zomatoDriver.find_elements_by_xpath(
    "//div[@class='grey-text nowrap fontsize5']")
# print(allAddress)
RestaurantAddress = []
for addressess in allAddress:
    RestaurantAddresses = addressess.text.strip()
    print(RestaurantAddresses)
    RestaurantAddress.append(RestaurantAddresses)
    # print(type(RestaurantAddress))

allOrderNumber = zomatoDriver.find_elements_by_xpath(
    "//div[@class='order-number']")
OrderNo = []
for orderNumber in allOrderNumber:
    OrderNos = orderNumber.text.strip()
    print(OrderNos)
    OrderNo.append(OrderNos)
    # print(type(OrderNo))

orderAmounts = zomatoDriver.find_elements_by_xpath("//div[@class='cost']")
OrderAmount = []
for orderAmount in orderAmounts:
    OrderAmounts = orderAmount.text.strip()
    print(OrderAmounts)
    OrderAmount.append(OrderAmounts)

print(OrderAmount)
j = 0
for i in OrderAmount:
    if i == ' ':
        i = i
    else:
        OrderAmount[j] = i.strip("₹")
    j += 1
print(OrderAmount)

for it in range(0, len(OrderAmount)):
    if OrderAmount[it] == '':
        OrderAmount[it] = ''
    elif ',' in OrderAmount[it]:
        OrderAmount[it] = OrderAmount[it].split(",")
        OrderAmount[it] = "".join(OrderAmount[it])
        OrderAmount[it] = float(OrderAmount[it])
    else:
        OrderAmount[it] = float(OrderAmount[it])
print(OrderAmount)

#totalAmountz = [x for x in totalAmount if isinstance(x, (int, float))]
totalAmountz = [x for x in OrderAmount if not isinstance(x, str)]
print(totalAmountz)

# print(type(OrderAmount))

datesOfOrdes = zomatoDriver.find_elements_by_xpath(
    "//div[@class='ui basic label']")
DatesOfOrder = []
for dateOfOrder in datesOfOrdes:
    DatesOfOrders = dateOfOrder.text.strip()
    print(DatesOfOrders)
    DatesOfOrder.append(DatesOfOrders)
print(DatesOfOrder)
# print(type(DatesOfOrders))

statusofOrder = zomatoDriver.find_elements_by_xpath(
    "//span[@class='right floated']")
OrderStatus = []
for orderStatus in statusofOrder:
    Order_Status = orderStatus.text.strip()
    print(Order_Status)
    OrderStatus.append(Order_Status)
    # print(type(OrderStatus))

fileName = 'ZomatoOrder.xlsx'
wb = Workbook()
ws1 = wb.active

ws1["A1"] = 'RestaurantName'
ws1["B1"] = 'RestaurantAddress'
ws1["C1"] = 'OrderNo'
ws1["D1"] = 'OrderAmount'
ws1["E1"] = 'DatesOfOrders'
ws1["F1"] = 'OrderStatus'

big_red_text = Font(color=colors.RED, size=14)

ws1["A1"].font = Font(bold=True)
ws1["B1"].font = Font(bold=True)
ws1["C1"].font = Font(bold=True)
ws1["D1"].font = Font(bold=True)
ws1["E1"].font = Font(bold=True)
ws1["F1"].font = Font(bold=True)
ws1["H1"].font = Font(bold=True)
ws1["H2"].font = Font(bold=True)

ws1["H1"].font = big_red_text
ws1["H2"].font = big_red_text

mapped = zip(RestaurantName, RestaurantAddress, OrderNo,
             OrderAmount, DatesOfOrder, OrderStatus)
mapped = set(mapped)
print("The mapped result is : ", end="")
print(mapped)
for row in mapped:
    ws1.append(row)

for orderTotal in totalAmountz:
    GrandTotal = sum(totalAmountz)

ws1["H1"] = 'Grand Total:'
ws1["H2"] = '₹' + str(GrandTotal)

wb.save('ZomatoOrder.xlsx')
print('We are done!')
