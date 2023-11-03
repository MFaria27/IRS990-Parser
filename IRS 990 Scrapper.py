from bs4 import BeautifulSoup
import urllib
import requests
import pandas as pd
import locale
import openpyxl

# Set locale to the US for currency formatting
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# To-Do List
# Finish Commenting Functions
# Better Way of searching for colleges
#   Probably going to be some form of web creation at this point
# Search "SCREAM" to find line 43
#   Some colleges have no tax fillings for the "college" but they do for things like "Alumni Association Inc"
#   Add break in code that returns (Could not find) if tax filings contain keywords like above

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

# Get all the information present on the IRS990 form on the occupations paid the greatest in the college
def get_institution_occupation_data(irs_990_soup):
    # We will be looking for these specific pieces of information
    # "Title Group" is the generalization of "Title" determined in a function below
    # "Total Compensation" is the addition of base comp and other comp
    columns = ["Name", "Title", "Title Group", "Base Compensation", "Other Comp", "Total Comp"]
    jobs = pd.DataFrame(columns=columns)
    company_wide_compensation = 0
    total_reported_employees = 0
    # SCREAM
    # Get the IRS990 part of the xml tax filling
    tax_filing_soup = irs_990_soup.find("IRS990")
    # PartVII in the tax filling is where all the occupational information is
    occupation_soup = tax_filing_soup.find_all("Form990PartVIISectionAGrp")
    for employee_soup in occupation_soup:
        name = ""
        # Forms have the name of an employee in different tasks, so try both
        try:
            name = employee_soup.find("PersonNm").text.strip()
        except:
            business_soup = employee_soup.find("BusinessName")
            name = business_soup.find("BusinessNameLine1Txt").text.strip()

        base_comp = int(employee_soup.find("ReportableCompFromOrgAmt").text.strip())
        other_comp = int(employee_soup.find("OtherCompensationAmt").text.strip())
        total_comp = base_comp + other_comp
        # Store the total compensation of every employee listed in the tax form
        company_wide_compensation += total_comp
        total_reported_employees += 1
        
        # If the employee is listed and doesn't make any money, skip them
        if total_comp == 0:
            continue

        title_name = employee_soup.find("TitleTxt").text.strip()
        # Use the title to get a generalization of the name 
        title_group = get_title_group(title_name)

        job_info = {
            "Name" : name,
            "Title": title_name,
            "Title Group" : title_group,
            "Base Compensation" : base_comp,
            "Other Comp" : other_comp,
            "Total Comp" : total_comp
        }
        # Add the employee's information to the list of jobs in the college
        jobs.loc[len(jobs)] = job_info

    # Sort all the employees by total compensation descending
    top_jobs = jobs.sort_values(by=['Total Comp'], ascending=False)
    return top_jobs, company_wide_compensation, total_reported_employees

# Use the title of an employee to generalize it into a group that can be shared by multiple
# "PRESIDENT (THROUGH 8/12)" and "PRESIDENT/CEO" get simplified into "President"
def get_title_group(title_name):
    # Name the base title group as other, so that if no title is found, report it as other
    tg = "Other"
    t = title_name.lower()
    # As of now, the code will look through this list going down. If it finds one title, it will skip the rest
    # If a title is "SECRETARY/VP/TREAS", it will be reported as "VP"
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

# Get the overarching revenue data of a college
def get_institution_revenue_data(irs_990_soup):
    beginYear = irs_990_soup.find("TaxPeriodBeginDt").text.strip()[:4]
    endYear = irs_990_soup.find("TaxPeriodEndDt").text.strip()[:4]
    year = beginYear + "-" + endYear

    filer = irs_990_soup.find("Filer")
    bName = filer.find("BusinessName")
    company = bName.find("BusinessNameLine1Txt").text.strip()

    pyTotalRevenue = int(irs_990_soup.find("PYTotalRevenueAmt").text.strip())
    cyTotalRevenue = int(irs_990_soup.find("CYTotalRevenueAmt").text.strip())

    pyNetRevenue = int(irs_990_soup.find("PYRevenuesLessExpensesAmt").text.strip())
    cyNetRevenue = int(irs_990_soup.find("CYRevenuesLessExpensesAmt").text.strip())

    revenue_dict = {
        "year" : year,
        "company" : company,
        "pyTotalRevenue" : pyTotalRevenue,
        "cyTotalRevenue" : cyTotalRevenue,
        "pyNetRevenue" : pyNetRevenue,
        "cyNetRevenue" : cyNetRevenue,
        "company_wide_compensation" : 0,
        "average_comp_per_reported" : 0,
        "net_over_comp_index" : 0
    }

    return revenue_dict

# Print the summary of the information found on a college (If asked for)
def institute_summary(revenue_dict, top_jobs):
    print(revenue_dict["company"])
    print(f"Current Year Total Revenue          :", locale.currency(revenue_dict["cyTotalRevenue"], grouping=True))
    print(f"Prior Year Total Revenue            :", locale.currency(revenue_dict["pyTotalRevenue"], grouping=True))
    print(f"Current Year Net Income Less Expense:", locale.currency(revenue_dict["cyNetRevenue"], grouping=True))
    print(f"Prior Year Net Income Less Expense  :", locale.currency(revenue_dict["pyNetRevenue"], grouping=True))
    print(f"Total College Wide Compensation     :", locale.currency(revenue_dict["company_wide_compensation"], grouping=True))
    print(top_jobs[0:10])

# Write all the scrapped information to an excel file
def write_intitution_to_excel(revenue_dict, top_jobs, nonprofit_subtitle, filename, index):
    
    # If a college or list of college excel file already exists for the list or single college, update the list instead of creating a new one
    try:
        excelWorkbook = openpyxl.load_workbook(filename)
    except:
        excelWorkbook = openpyxl.Workbook()

    # If a college in the excel file already exists, overwrite the sheet instead of making a new one
    try:
        sheet = excelWorkbook[nonprofit_subtitle]
    except:
        excelWorkbook.create_sheet(title=nonprofit_subtitle)
        sheet = excelWorkbook[nonprofit_subtitle]
    sheet["A1"] = "Year Range"
    sheet["B1"] = revenue_dict["year"]
    sheet["A2"] = "Current Year Total Revenue"
    sheet["B2"] = locale.currency(revenue_dict["cyTotalRevenue"], grouping=True)
    sheet["A3"] = "Prior Year Total Revenue"
    sheet["B3"] = locale.currency(revenue_dict["pyTotalRevenue"], grouping=True)
    sheet["A4"] = "Current Year Net Income Less Expense"
    sheet["B4"] = locale.currency(revenue_dict["cyNetRevenue"], grouping=True)
    sheet["A5"] = "Prior Year Net Income Less Expense"
    sheet["B5"] = locale.currency(revenue_dict["pyNetRevenue"], grouping=True)
    sheet["A6"] = "Total College Wide Compensation"
    sheet["B6"] = locale.currency(revenue_dict["company_wide_compensation"], grouping=True)
    sheet["A7"] = "Average Comp Per Reported Employee"
    sheet["B7"] = locale.currency(revenue_dict["average_comp_per_reported"], grouping=True)
    sheet["A8"] = "Net Income / Total Comp Index"
    sheet["B8"] = revenue_dict["net_over_comp_index"]

    # Create a dictionary of the title groups to count how many times each group shows up in the school irs form
    num_of_titles = {
        "Vice President" : 0 ,
        "Vice Provost" : 0,
        "President" : 0, 
        "Provost" : 0, 
        "VP" : 0, 
        "Trustee" : 0,
        "Dean" : 0, 
        "Exec" : 0, 
        "Prof" : 0,
        "Treas" : 0,
        "Secretary" : 0,
        "Chief" : 0,
        "Dept Head" : 0,
        "Other" : 0
    }

    sheet["B10"] = "Name"
    sheet["C10"] = "Title"
    sheet["D10"] = "Title Group"
    sheet["E10"] = "Base Compensation"
    sheet["F10"] = "Other Comp"
    sheet["G10"] = "Total Comp"

    row = 11
    job_index = 1
    for i,job in top_jobs.iterrows():
        
        sheet["A"+str(row)] = job_index
        sheet["B"+str(row)] = job["Name"]
        sheet["C"+str(row)] = job["Title"]
        sheet["D"+str(row)] = job["Title Group"]
        num_of_titles[job["Title Group"]] = num_of_titles.get(job["Title Group"]) + 1
        sheet["E"+str(row)] = job["Base Compensation"]
        sheet["F"+str(row)] = job["Other Comp"]
        sheet["G"+str(row)] = job["Total Comp"]
        row += 1
        job_index += 1

    # Add college summary information to the Master Sheet
    sheet = excelWorkbook["Sheet"]
    sheet["A1"] = "Year"
    sheet["B1"] = "School"
    sheet["C1"] = "Revenue"
    sheet["D1"] = "Net Income"
    sheet["E1"] = "Total Employee Comp"
    sheet["F1"] = "Average Comp"
    sheet["G1"] = "Net/Comp Index"
    sheet["H1"] = "Presidents"
    sheet["I1"] = "Vice Presidents"
    sheet["J1"] = "Provosts"
    sheet["K1"] = "Trustees"
    sheet["L1"] = "Deans"
    sheet["M1"] = "Executives"
    sheet["N1"] = "Professors"
    sheet["O1"] = "Treasurers"
    sheet["P1"] = "Secretaries"
    sheet["Q1"] = "Chiefs"
    sheet["R1"] = "Dept Heads"
    sheet["S1"] = "Other"

    summary_index = int((index / 2) + 2)

    sheet["A" + str(summary_index)] = revenue_dict["year"]
    sheet["B" + str(summary_index)] = nonprofit_subtitle
    sheet["C" + str(summary_index)] = locale.currency(revenue_dict["cyTotalRevenue"], grouping=True)
    sheet["D" + str(summary_index)] = locale.currency(revenue_dict["cyNetRevenue"], grouping=True)
    sheet["E" + str(summary_index)] = locale.currency(revenue_dict["company_wide_compensation"], grouping=True)
    sheet["F" + str(summary_index)] = locale.currency(revenue_dict["average_comp_per_reported"], grouping=True)
    sheet["G" + str(summary_index)] = revenue_dict["net_over_comp_index"]
    sheet["H" + str(summary_index)] = num_of_titles["President"]
    sheet["I" + str(summary_index)] = num_of_titles["Vice President"] + num_of_titles["VP"]
    sheet["J" + str(summary_index)] = num_of_titles["Provost"] + num_of_titles["Vice Provost"]
    sheet["K" + str(summary_index)] = num_of_titles["Trustee"]
    sheet["L" + str(summary_index)] = num_of_titles["Dean"]
    sheet["M" + str(summary_index)] = num_of_titles["Exec"]
    sheet["N" + str(summary_index)] = num_of_titles["Prof"]
    sheet["O" + str(summary_index)] = num_of_titles["Treas"]
    sheet["P" + str(summary_index)] = num_of_titles["Secretary"]
    sheet["Q" + str(summary_index)] = num_of_titles["Chief"]
    sheet["R" + str(summary_index)] = num_of_titles["Dept Head"]
    sheet["S" + str(summary_index)] = num_of_titles["Other"]

    # Save the created workbook into the filename provided
    excelWorkbook.save(filename)

# "Main" function. Searches for a college or list of colleges
def search(nonprofit_search_query, nonprofit_subtitle, show_summary, filename, index):
    # Use ProPublica API to get basic institution data
    search_url = "https://projects.propublica.org/nonprofits/api/v2/search.json?"
    parameters = urllib.parse.urlencode({
        'q' : nonprofit_search_query
    })
    search_result = requests.get(search_url+parameters)

    # If an institution is found, continue; break if nothing is found
    if search_result.status_code == 200:
        print("Found", nonprofit_search_query)

        # Get the web content in the form of an xml version of an IRS 990 form
        irs_990_soup = get_irs_990_web_content(search_result)

        # Use xml content to get the paid reported occupations sorted by highest compensation descending
        top_jobs, company_wide_compensation, total_reported_employees = get_institution_occupation_data(irs_990_soup)

        # Get the basic IRS data from the xml content
        revenue_dict = get_institution_revenue_data(irs_990_soup)
        revenue_dict["company_wide_compensation"] = company_wide_compensation
        try:
            revenue_dict["average_comp_per_reported"] = company_wide_compensation / total_reported_employees
        except:
            revenue_dict["average_comp_per_reported"] = 0
        try:
            revenue_dict["net_over_comp_index"] = revenue_dict["cyNetRevenue"] / company_wide_compensation
        except:
            revenue_dict["net_over_comp_index"] = 0
        
        # If you would like to see a terminal summary of the information found, add "True" to the function parameters
        # Default is that a summary will be shown if looking for a single college, and will not be shown for a list search
        if(show_summary):
            institute_summary(revenue_dict, top_jobs)

        # Write all the information to the Master Excel file
        write_intitution_to_excel(revenue_dict, top_jobs, nonprofit_subtitle, filename, index)
    
    else:
        # If the college is not found, report it
        print("Could not find", nonprofit_search_query)


print("If reading from list, enter the filename. If not, just hit enter.")
using_list = input("Using list?: ")
# Make sure that if a list is used, that is is present in a .txt file format
# See "Institutes.txt" for an example
if len(using_list) > 4 and using_list[-4:] == ".txt":
    print("Reading from file", using_list, "...")
    institutes = []
    # Get the list of institutes and institute subtitles
    with open("Inputs/"+using_list, "r") as f:
        institutes = f.readlines()
    # Every line has a "/n" indicating a new line in the text file, so get rid of them
    for i in range(len(institutes)):
        institutes[i] = institutes[i][:-1]
    
    # Look through every even numbered line for the full institute name to be searched (Odd is the subtitle)
    for i in range(len(institutes)):
        if i % 2 == 0:
            # Every college will be saved in a sheet in a master excel file, named after the name of the txt file the list is in
            # Ex. Using the lists of colleges in "Institutes.txt" will output in "Institutes.xlsx"
            search(institutes[i], institutes[i+1], False, "Outputs/"+using_list[:-4]+".xlsx",i)
        else:
            continue
    print("Output college information in", "Outputs/"+using_list[:-4]+".xlsx")

else:
    # Ask for Institution Name + Institution Subtitle (Excel Sheet Name)
    nonprofit_search_query = input("Full Institute Name: ")
    nonprofit_subtitle = input("Institute Nickname: ")
    # If no subtitle is provided, just make it the search query
    if nonprofit_subtitle == "":
        nonprofit_subtitle = nonprofit_search_query
    # The output file for a single query will be stored in the input subtitles name + .xlsx
    #Ex. If I look up Worcester Polytechnic Institute (WPI), it will be saved as WPI.xlsx
    search(nonprofit_search_query, nonprofit_subtitle, True, "Outputs/"+nonprofit_subtitle+".xlsx",1)
    print("Output college information in", "Outputs/"+nonprofit_subtitle+".xlsx")

