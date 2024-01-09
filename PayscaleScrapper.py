from bs4 import BeautifulSoup
import requests
import pandas as pd

columns = ["Page", "Rank", "School Name", "20 Year Net ROI", "Total 4 Year Cost", "Graduation Rate", "Typical Years to Graduate", "Average Loan Amount"]
college_roi = pd.DataFrame(columns=columns)

page_url = "https://www.payscale.com/college-roi/page/"
pages_available = 153

for page in range(1,pages_available + 1):

    page_result = requests.get(page_url+str(page))
    page_content = page_result.content
    page_soup = BeautifulSoup(page_content, "html.parser")
    row_soup = page_soup.find_all("tr", class_="data-table__row")

    for row in row_soup:
        row_values = [page]
        cell_soup = row.find_all("td", class_="data-table__cell")
        for cell in cell_soup:
            row_values.append(cell.find("span", class_="data-table__value").text.strip())
        college_roi.loc[len(college_roi)] = row_values

print(college_roi)
college_roi.to_excel('Outputs/Payscale.xlsx')
    

