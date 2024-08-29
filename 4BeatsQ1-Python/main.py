import selenium
import openpyxl
import datetime
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By


current_datetime = datetime.datetime.now()
current_day = current_datetime.strftime("%A")
print(f"Current Day:{current_day}")


filename = r"4BeatsQ1.xlsx"
df = load_workbook(filename)
sheet = df[current_day]

# Needs to be same as number ok keywords present in Excel Sheet
num_keywords = 10

for row in range(3, 3+num_keywords):
    keyword = sheet['C'+str(row)].value
    print(f"Searching for keyword:{keyword}")

    driver = webdriver.Chrome()
    driver.get("https://www.google.com/")
    driver.find_element(By.NAME, "q").send_keys(keyword)
    time.sleep(2)

    options = driver.find_elements(By.CLASS_NAME, "lnnVSe")
    text = options[0].get_attribute('aria-label')

    longest = text
    len_longest = len(text)
    shortest = text
    len_shortest = len(shortest)

    for option in options[1:10]:
        text = option.get_attribute('aria-label')

        # If there are multiple options with the same length, they will ALL be saved

        if len(text) > len_longest:
            longest = text
            len_longest = len(text)

        elif len(text) == len(longest):
            longest += ", " + text

        elif len(text) < len_shortest:
            shortest = text
            len_shortest = len(text)

        elif len(text) == len(shortest):
            shortest += ", " + text

    sheet['D'+str(row)] = longest
    sheet['E'+str(row)] = shortest
    driver.close()

    print(f"Longest search: {longest}")
    print(f"Shortest search: {shortest}")
    print()


df.save(filename)
print("Keyword searching- Completed")

