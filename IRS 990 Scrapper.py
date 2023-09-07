from bs4 import BeautifulSoup
import urllib
import requests
import pandas as pd
from pprint import pprint
import locale
from openpyxl import load_workbook

locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# To-Do List
# Create a library of checked colleges

def get_irs_990_web_content(search_result):
    search_result_json = search_result.json()
    search_result_ein = search_result_json['organizations'][0]['ein']

    specific_nonprofit_search = "https://projects.propublica.org/nonprofits/organizations/" + str(search_result_ein)
    page = requests.get(specific_nonprofit_search).content

    page_soup = BeautifulSoup(page, "html.parser")
    xml_soup = page_soup.find(class_="action xml")
    xml_url_query = xml_soup.get("href")
    xml_url = "https://projects.propublica.org" + xml_url_query

    xml_search_result = requests.get(xml_url).content

    irs_990_soup = BeautifulSoup(xml_search_result, features="xml")
    return irs_990_soup

def get_institution_occupation_data(irs_990_soup):
    columns = ["Name", "Title", "Base Compensation", "Other Comp", "Total Comp"]
    jobs = pd.DataFrame(columns=columns)
    company_wide_compensation = 0
    tax_filing_soup = irs_990_soup.find("IRS990")
    occupation_soup = tax_filing_soup.find_all("Form990PartVIISectionAGrp")
    for employee_soup in occupation_soup:
        name = ""
        try:
            name = employee_soup.find("PersonNm").text.strip()
        except:
            business_soup = employee_soup.find("BusinessName")
            name = business_soup.find("BusinessNameLine1Txt").text.strip()

        base_comp = int(employee_soup.find("ReportableCompFromOrgAmt").text.strip())
        other_comp = int(employee_soup.find("OtherCompensationAmt").text.strip())
        total_comp = base_comp + other_comp
        company_wide_compensation += total_comp
        
        if total_comp == 0:
            continue

        job_info = {
            "Name" : name,
            "Title": employee_soup.find("TitleTxt").text.strip(),
            "Base Compensation" : base_comp,
            "Other Comp" : other_comp,
            "Total Comp" : total_comp
        }
        jobs.loc[len(jobs)] = job_info

    top_jobs = jobs.sort_values(by=['Total Comp'], ascending=False)
    return top_jobs, company_wide_compensation

def get_institution_revenue_data(irs_990_soup):
    filer = irs_990_soup.find("Filer")
    bName = filer.find("BusinessName")
    company = bName.find("BusinessNameLine1Txt").text.strip()

    pyTotalRevenue = int(irs_990_soup.find("PYTotalRevenueAmt").text.strip())
    cyTotalRevenue = int(irs_990_soup.find("CYTotalRevenueAmt").text.strip())

    pyNetRevenue = int(irs_990_soup.find("PYRevenuesLessExpensesAmt").text.strip())
    cyNetRevenue = int(irs_990_soup.find("CYRevenuesLessExpensesAmt").text.strip())

    revenue_dict = {
        "company" : company,
        "pyTotalRevenue" : pyTotalRevenue,
        "cyTotalRevenue" : cyTotalRevenue,
        "pyNetRevenue" : pyNetRevenue,
        "cyNetRevenue" : cyNetRevenue,
        "company_wide_compensation" : 0
    }

    return revenue_dict

def institute_summary(revenue_dict, top_jobs):
    print(revenue_dict["company"])
    print(f"Current Year Total Revenue          :", locale.currency(revenue_dict["cyTotalRevenue"], grouping=True))
    print(f"Prior Year Total Revenue            :", locale.currency(revenue_dict["pyTotalRevenue"], grouping=True))
    print(f"Current Year Net Income Less Expense:", locale.currency(revenue_dict["cyNetRevenue"], grouping=True))
    print(f"Prior Year Net Income Less Expense  :", locale.currency(revenue_dict["pyNetRevenue"], grouping=True))
    print(f"Total Company Wide Compensation     :", locale.currency(revenue_dict["company_wide_compensation"], grouping=True))
    print(top_jobs[0:10])

def write_intitution_to_excel(revenue_dict, top_jobs):
    excelWorkbook = load_workbook("Master.xlsx")
    writer = pd.ExcelWriter("Master.xlsx", engine='openpyxl')
    writer.book = excelWorkbook
    try:
        sheet = excelWorkbook[nonprofit_subtitle]
    except:
        excelWorkbook.create_sheet(title=nonprofit_subtitle)
        sheet = excelWorkbook[nonprofit_subtitle]
    sheet["A1"] = "Current Year Total Revenue"
    sheet["B1"] = locale.currency(revenue_dict["cyTotalRevenue"], grouping=True)
    sheet["A2"] = "Prior Year Total Revenue"
    sheet["B2"] = locale.currency(revenue_dict["pyTotalRevenue"], grouping=True)
    sheet["A3"] = "Current Year Net Income Less Expense"
    sheet["B3"] = locale.currency(revenue_dict["cyNetRevenue"], grouping=True)
    sheet["A4"] = "Prior Year Net Income Less Expense"
    sheet["B4"] = locale.currency(revenue_dict["pyNetRevenue"], grouping=True)
    sheet["A5"] = "Total Company Wide Compensation"
    sheet["B5"] = locale.currency(revenue_dict["company_wide_compensation"], grouping=True)

    sheet["B7"] = "Name"
    sheet["C7"] = "Title"
    sheet["D7"] = "Base Compensation"
    sheet["E7"] = "Other Comp"
    sheet["F7"] = "Total Comp"

    row = 8
    index = 1
    for i,job in top_jobs.iterrows():
        
        sheet["A"+str(row)] = index
        sheet["B"+str(row)] = job["Name"]
        sheet["C"+str(row)] = job["Title"]
        sheet["D"+str(row)] = job["Base Compensation"]
        sheet["E"+str(row)] = job["Other Comp"]
        sheet["F"+str(row)] = job["Total Comp"]
        row += 1
        index += 1

    writer.save()
    writer.close()


nonprofit_search_query = input("Full Institute Name: ")
nonprofit_subtitle = input("Institute Nickname: ")
if nonprofit_subtitle == "":
    nonprofit_subtitle = nonprofit_search_query
# Stevens Institute of Technology
# Worcester Polytechnic Institute
# Rochester Institute of Technology

search_url = "https://projects.propublica.org/nonprofits/api/v2/search.json?"
parameters = urllib.parse.urlencode({
      'q' : nonprofit_search_query
  })

search_result = requests.get(search_url+parameters)
if search_result.status_code == 200:
    irs_990_soup = get_irs_990_web_content(search_result)

    top_jobs, company_wide_compensation = get_institution_occupation_data(irs_990_soup)

    revenue_dict = get_institution_revenue_data(irs_990_soup)
    revenue_dict["company_wide_compensation"] = company_wide_compensation

    institute_summary(revenue_dict, top_jobs)

    write_intitution_to_excel(revenue_dict, top_jobs)
    
else:
    print("Could not find", nonprofit_search_query)

