import openpyxl
import datetime
from openpyxl.styles import Alignment
from openpyxl.styles import Font, Fill
from selenium import webdriver
from PIL import Image
# give access to enter key or escape key
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import json
import os.path

x = datetime.datetime.now()
out = datetime.datetime(2021, 2, 28)


PATH = "C:\Program Files (x86)\chromedriver.exe"

options = webdriver.ChromeOptions()
options.add_argument('--lang=es')
options.add_argument('--ignore-certificate-errors')
options.add_argument("--test-type")

driver = webdriver.Chrome(executable_path=PATH, options=options)

link = ""
linkFile = os.path.isfile("./link.txt")


if linkFile:
    j = open("./link.txt", "r")

    data = j.read()
    print(data)
    time.sleep(2)
    link = str(data)

if x.day >= out.day:
    print("Expired")
else:
    # place where we scrap the data
    driver.get(link)


# wait until load's correctly
time.sleep(30)

# specify how much reviews you want to save
# first review index
startRevNum = 0
# last review index
endRevNum = 5

reviewRangeFile = os.path.isfile("./reviewRange.txt")
if reviewRangeFile:
    f = open("./reviewRange.txt")
    data = json.load(f)
    time.sleep(2)
    startRevNum = data[0]
    endRevNum = data[1]

else:
    # create blank file
    f = open("./reviewRange.txt", "w+")
    rangeList = [0, 5]
    json.dump(rangeList, f)

excelPath = "./reviews.xlsx"

time.sleep(2)
workBook = openpyxl.load_workbook(excelPath)
sheet = workBook.active

# index used to store review picture
i = 0

# find 'sort' dropdown
dropdown = driver.find_element_by_class_name("dkSGpd")
# click in dropdown
dropdown.click()
# wait to load dropdown options
time.sleep(2)
# get dropdown elements
d = driver.find_elements_by_class_name("zZoSGe")
# click on 'most recent' element
d[1].click()
# wait until action ends
time.sleep(1)

# create action chain object used to simulate keystroke
action = ActionChains(driver)

# get first comment
comment = driver.find_element_by_class_name("PuaHbe")
# focus first comment
comment.click()
# load next block of comments
# each block contains 10 comments

nameList = []
# count occuped rows
val = 2
while sheet['A' + str(val)].value is not None:
    nameList.append(sheet['A' + str(val)].value)
    val += 1

i = val - 2

# preload some reviews block
oldReviewsAmount = len(driver.find_elements_by_xpath('//div[@jscontroller="e6Mltc"]'))
comment.send_keys(Keys.END)
# sleep until next block of comments is loaded
time.sleep(3)
newReviewsAmount = len(driver.find_elements_by_xpath('//div[@jscontroller="e6Mltc"]'))
reviews = driver.find_elements_by_xpath('//div[@class="jxjCjc"]')

# load reviews block until revNum is reached or no more reviews avaiable to load
while (oldReviewsAmount < newReviewsAmount) and (endRevNum > newReviewsAmount):
    oldReviewsAmount = len(driver.find_elements_by_xpath('//div[@jscontroller="e6Mltc"]'))
    comment.send_keys(Keys.END)
    # sleep until next block of comments is loaded
    time.sleep(3)
    newReviewsAmount = len(driver.find_elements_by_xpath('//div[@jscontroller="e6Mltc"]'))

# update reviews container with loaded reviews
reviewsContainer = driver.find_elements_by_id("reviewSort")
# get all reviews loaded
reviews = driver.find_elements_by_xpath('//div[@class="jxjCjc"]')

time.sleep(1)

# get data from reviews
for e in reviews[startRevNum:endRevNum + 1]:
    i += 1
    # get reviewer name
    nameDiv = e.find_element_by_class_name("TSUbDb")
    reviewerName = nameDiv.find_element_by_tag_name("a").text
    reviewerCommentContainer = e.find_element_by_class_name("Jtu6Td")
    reviewerComment = reviewerCommentContainer.find_element_by_tag_name("span").text

    cell = sheet.cell(row=1 + i, column=1)
    cell.value = reviewerName
    cell.font = Font(size=36)
    cell.alignment = Alignment(horizontal='center')
    cell.alignment = Alignment(vertical='center')

    # go to element before taking a screenshoot to prevent
    # capture bugs
    #action.move_to_element(e).perform()
    driver.execute_script("arguments[0].focus()", nameDiv.find_element_by_tag_name("a"))
    # create empty image
    img = Image.new("RGB", (300, 300), (0, 0, 0))
    imageName = str("./screenshot/review" + str(i) + ".png")
    img.save(imageName, "PNG")

    # take screenshot of element and store in file
    e.screenshot("./screenshot/review" + str(i) + ".png")

    imagen = openpyxl.drawing.image.Image("./screenshot/review" + str(i) + ".png")
    sheet.add_image(imagen, 'B' + str(1 + i))
    # save it in a list

workBook.save(excelPath)
time.sleep(3)
driver.close()