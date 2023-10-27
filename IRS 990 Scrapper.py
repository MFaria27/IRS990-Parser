from bs4 import BeautifulSoup
import urllib
import requests
import pandas as pd
import locale
from openpyxl import load_workbook

# Set locale to the US for currency formatting
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# To-Do List
# Finish Commenting Functions
# Fix Excel Corruption
# For some reason, title group "President" cannot be read
# Better Way of searching for colleges
#   Probably going to be some form of web creation at this point

# Use the information from ProPublica API and more GET requests to get the xml version of the IRS 990 form
def get_irs_990_web_content(search_result):

    # Get the identification number for the institution to be used in another url
    search_result_json = search_result.json()
    search_result_ein = search_result_json['organizations'][0]['ein']

    # Search for the base html page of the searched Non-profit 
    specific_nonprofit_search = "https://projects.propublica.org/nonprofits/organizations/" + str(search_result_ein)
    page = requests.get(specific_nonprofit_search).content

    # Use BeautifulSoup library to find the first xml button on the website and get the link to the xml report
    page_soup = BeautifulSoup(page, "html.parser")
    xml_soup = page_soup.find_all(lambda tag: tag.name == "a" and "XML" in tag.text)
    xml_url_query = xml_soup[0].get("href")
    xml_url = "https://projects.propublica.org" + xml_url_query
    xml_search_result = requests.get(xml_url).content

    # Use BeautifulSoup to parse the web content as xml and return
    irs_990_soup = BeautifulSoup(xml_search_result, features="xml")
    return irs_990_soup

def get_institution_occupation_data(irs_990_soup):
    columns = ["Name", "Title", "Title Group", "Base Compensation", "Other Comp", "Total Comp"]
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

        title_name = employee_soup.find("TitleTxt").text.strip()
        title_group = get_title_group(title_name)

        job_info = {
            "Name" : name,
            "Title": title_name,
            "Title Group" : title_group,
            "Base Compensation" : base_comp,
            "Other Comp" : other_comp,
            "Total Comp" : total_comp
        }
        jobs.loc[len(jobs)] = job_info

    top_jobs = jobs.sort_values(by=['Total Comp'], ascending=False)
    return top_jobs, company_wide_compensation

def get_title_group(title_name):
    tg = "Other"
    t = title_name.lower()
    list_of_titles = [
        "Vice President",
        "Vice Provost",
        "President", 
        "Provost", 
        "VP", 
        "Trustee",
        "Dean", 
        "Exec", 
        "Prof",
        "Treas",
        "Secretary",
        "Chief",
        "Dept Head"
    ]
    
    for title in list_of_titles:
        if title.lower() in t:
            tg = title
            break
    return tg

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

def write_intitution_to_excel(revenue_dict, top_jobs, nonprofit_subtitle):
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
    sheet["D7"] = "Title Group"
    sheet["E7"] = "Base Compensation"
    sheet["F7"] = "Other Comp"
    sheet["G7"] = "Total Comp"

    row = 8
    index = 1
    for i,job in top_jobs.iterrows():
        
        sheet["A"+str(row)] = index
        sheet["B"+str(row)] = job["Name"]
        sheet["C"+str(row)] = job["Title"]
        sheet["D"+str(row)] = job["Title Group"]
        sheet["E"+str(row)] = job["Base Compensation"]
        sheet["F"+str(row)] = job["Other Comp"]
        sheet["G"+str(row)] = job["Total Comp"]
        row += 1
        index += 1

    writer.save()
    writer.close()

def search(nonprofit_search_query, nonprofit_subtitle, show_summary):
    # Use ProPublica API to get basic institution data
    search_url = "https://projects.propublica.org/nonprofits/api/v2/search.json?"
    parameters = urllib.parse.urlencode({
        'q' : nonprofit_search_query
    })
    search_result = requests.get(search_url+parameters)

    # If an institution is found, continue; break if nothing is found
    if search_result.status_code == 200:

        # Get the web content in the form of an xml version of an IRS 990 form
        irs_990_soup = get_irs_990_web_content(search_result)

        # Use xml content to get the paid reported occupations sorted by highest compensation descending
        top_jobs, company_wide_compensation = get_institution_occupation_data(irs_990_soup)

        # Get the basic IRS data from the xml content
        revenue_dict = get_institution_revenue_data(irs_990_soup)
        revenue_dict["company_wide_compensation"] = company_wide_compensation

        # Erase Comment to get a terminal summary of the information found
        if(show_summary):
            institute_summary(revenue_dict, top_jobs)

        # Write all the information to the Master Excel file
        write_intitution_to_excel(revenue_dict, top_jobs, nonprofit_subtitle)
    
    else:
        print("Could not find", nonprofit_search_query)


print("If reading from list, enter 'list'. If not, just hit enter.")
using_list = input("Using list?: ")
if using_list == "list":
    print("Reading from file...")
    institutes = []
    with open("Institutes.txt", "r") as f:
        institutes = f.readlines()
    for i in range(len(institutes)):
        institutes[i] = institutes[i][:-1]
    for i in range(len(institutes)):
        if i % 2 == 0:
            search(institutes[i], institutes[i+1], False)
        else:
            continue

else:
    # Ask for Institution Name + Excel Sheet Name
    nonprofit_search_query = input("Full Institute Name: ")
    nonprofit_subtitle = input("Institute Nickname: ")
    # If no subtitle is provided, just make it the search query
    if nonprofit_subtitle == "":
        nonprofit_subtitle = nonprofit_search_query
    search(nonprofit_search_query, nonprofit_subtitle, True)

