from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from xlwt import Workbook

PATH = "C:\Program Files (x86)\chromedriver.exe"

driver = webdriver.Chrome(PATH)

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet("SSPN-2017")

# Adding the headings
head = [
    "ROLL NO.",
    "NAME",
    "MOTHERS'S NAME",
    "FATHER'S NAME",
    "ENGLISH-301",
    "MATHEMATICS-041",
    "PHYSICS(T)-042",
    "PHYSICS(P)-042",
    "CHEMISTRY(T)-043",
    "CHEMISTRY(P)-043",
    "BIOLOGY/COMPUTER SCIENCE(T)",
    "BIOLOGY/COMPUTER SCIENCE(P)",
    "WORK EXPERIENCE(500)",
    "PHY & HEALTH EDUCA(502)",
    "GENERAL STUDIES(503)",
]

for m in range(len(head)):
    sheet1.write(1, 1 + m, head[m])


SCHOOL_NUMBER = 32518
CENTER_NUMBER = 3267
rno = list(range(3641848, 3641889))  # Roll numbers

for count,roll in enumerate(rno):  # Roll numbers:
    # go to result website
    driver.get(
        "https://resultsarchives.nic.in/cbseresults/cbseresults2017/class12npy/Class12th17.htm"
    )

    # entering the roll numbers
    search = driver.find_element("name", "regno")
    search.send_keys(roll)
    search = driver.find_element("name", "sch")
    search.send_keys(SCHOOL_NUMBER)
    search = driver.find_element("name", "cno")
    search.send_keys(CENTER_NUMBER)
    search.send_keys(Keys.RETURN)

    name = driver.find_element("xpath",
        "/html/body/div/table[1]/tbody/tr[2]/td[2]/font/b"
    ).text
    mname = driver.find_element("xpath",
        "/html/body/div/table[1]/tbody/tr[3]/td[2]/font"
    ).text
    fname = driver.find_element("xpath",
        "/html/body/div/table[1]/tbody/tr[4]/td[2]/font"
    ).text
    eng = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[2]/td[3]/font"
        ).text[1:]
    )
    math = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[3]/td[3]/font"
        ).text[1:]
    )
    phy = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[4]/td[3]/font"
        ).text[1:]
    )
    phyp = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[4]/td[4]/font"
        ).text[1:]
    )
    chem = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[5]/td[3]/font"
        ).text[1:]
    )
    chemp = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[5]/td[4]/font"
        ).text[1:]
    )
    bio = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[6]/td[3]/font"
        ).text[1:]
    )
    biop = int(
        driver.find_element("xpath",
            "/html/body/div/div/center/table/tbody/tr[6]/td[4]/font"
        ).text[1:]
    )
    we = driver.find_element("xpath",
        "/html/body/div/div/center/table/tbody/tr[7]/td[6]/font"
    ).text
    pe = driver.find_element("xpath",
        "/html/body/div/div/center/table/tbody/tr[8]/td[6]/font"
    ).text
    gs = driver.find_element("xpath",
        "/html/body/div/div/center/table/tbody/tr[9]/td[6]/font"
    ).text

    output = [
        roll,
        name,
        mname,
        fname,
        eng,
        math,
        phy,
        phyp,
        chem,
        chemp,
        bio,
        biop,
        we,
        pe,
        gs,
    ]


    # writing to results.xls
    # for l in range(len(output)):
    #     sheet1.write(1 + count, 1 + l, output[l])

    print(output)

# Saving the workbook
wb.save("result.xls")

print("workbook saved")
driver.close()
