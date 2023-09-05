from bs4 import BeautifulSoup
import urllib
import requests
import pandas as pd
from pprint import pprint

columns = ["Name", "Title", "Compensation", "Other Comp"]
jobs = pd.DataFrame(columns=columns)

nonprofit_search_query = "Stevens Institute of Technology"

search_url = "https://projects.propublica.org/nonprofits/api/v2/search.json?"
parameters = urllib.parse.urlencode({
      'q' : nonprofit_search_query
  })

search_result = requests.get(search_url+parameters)
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
filer = irs_990_soup.find("Filer")
bName = filer.find("BusinessName")
company = bName.find("BusinessNameLine1Txt").text.strip()

print(company)

irs990_soup = irs_990_soup.find("IRS990")
occupation_soup = irs990_soup.find_all("Form990PartVIISectionAGrp")
for employee_soup in occupation_soup:
    job_info = {
        "Name" : employee_soup.find("PersonNm").text.strip(),
        "Title": employee_soup.find("TitleTxt").text.strip(),
        "Compensation" : int(employee_soup.find("ReportableCompFromOrgAmt").text.strip()),
        "Other Comp" : int(employee_soup.find("OtherCompensationAmt").text.strip())
    }
    jobs.loc[len(jobs)] = job_info

top_jobs = jobs.sort_values(by=['Compensation'], ascending=False)[0:10]

print(top_jobs)
#top_jobs.to_csv("top_jobs.csv")
    


