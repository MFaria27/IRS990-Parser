from bs4 import BeautifulSoup
import requests

search_url = "https://projects.propublica.org/nonprofits/api/v2/search.json?q=Worcester+Polytechnic+Institute"

search_result = requests.get(search_url)
search_result_json = search_result.json()
search_result_ein = search_result_json['organizations'][0]['ein']

specific_nonprofit_search = "https://projects.propublica.org/nonprofits/organizations/" + str(search_result_ein)
page = requests.get(specific_nonprofit_search).content

page_soup = BeautifulSoup(page, "html.parser")
xml_soup = page_soup.find(class_="action xml")
xml_url_query = xml_soup.get("href")
xml_url = "https://projects.propublica.org" + xml_url_query

xml_search_result = requests.get(xml_url)
print(xml_search_result.content)