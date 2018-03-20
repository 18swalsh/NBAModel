from selenium import webdriver
import xlsxwriter
import time

start_time = time.clock()

# set up webdriver
browser = webdriver.Chrome(executable_path="C:/Users/cwalsh/Documents/Steve/chromedriver/chromedriver")

# create excel workbook
workbook = xlsxwriter.Workbook("C:/Users/cwalsh/Documents/Steve/NBA_Output_"+str(round(time.time(), 2))+".xlsx")


def get_values(url, sheet_name):
    browser.get(url)
    table_headers = browser.find_elements_by_xpath("//table/thead/tr/th")
    headers = [x.text for x in table_headers]
    header_count = sum(1 for i in headers if i is not "")
    data_table = browser.find_elements_by_xpath("//table/tbody/tr/td")
    values = [x.text for x in data_table]

    print("Values retrieved for " + sheet_name)

    # create & activate sheet
    workbook.add_worksheet(sheet_name)
    worksheet = workbook.get_worksheet_by_name(sheet_name)

    # write values
    for y in range(0, header_count):
        worksheet.write(0, y, headers[y])

    index = 0
    for z in range(1, 31):  # 30 NBA Teams
        for x in range(0, header_count):
            worksheet.write(z, x, values[index])
            index += 1


# call get_values for each url
get_values("https://stats.nba.com/teams/traditional/?sort=W&dir=-1", "Teams General Traditional")
get_values("https://stats.nba.com/teams/advanced/?sort=W&dir=-1", "Teams General Advanced")
get_values("https://stats.nba.com/teams/four-factors/?sort=W&dir=-1", "Teams General Four Factors")
get_values("https://stats.nba.com/teams/misc/?sort=W&dir=-1", "Teams General Misc")
get_values("https://stats.nba.com/teams/scoring/?sort=W&dir=-1", "Teams General Scoring")
get_values("https://stats.nba.com/teams/opponent/?sort=W&dir=-1", "Teams General Opponent")
get_values("https://stats.nba.com/teams/defense/?sort=W&dir=-1", "Teams General Defense")

# close workbook and browser
workbook.close()
browser.quit()

print("Model created in " + str(round((time.clock() - start_time)/60, 2)) + " minutes.")
