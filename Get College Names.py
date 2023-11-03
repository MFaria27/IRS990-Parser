import urllib
import requests
import locale
import pprint

# This code gets every nonprofit organization with the ntee code '2' (Education/Educational Services) and the keywork "University"
# All of the results get placed into "Colleges.txt" 
# Note: May take time to run

#Number of pages: 400
with open("college names.txt", "w") as colleges:
    for page_index in range(0,400):
        colleges.write("Page " + str(page_index) + "\n")
        search_url = "https://projects.propublica.org/nonprofits/api/v2/search.json?"
        parameters = urllib.parse.urlencode({
            'q' : "University",
            'ntee%5Bid%5D' : 2,
            'page' : page_index
        })
        search_result = requests.get(search_url+parameters)
        search_json = search_result.json()
        for college in range(0,25):
            colleges.write(search_json['organizations'][college]['name'] + "\n")
        if page_index % 20 ==0: 
            print("Completed " +str(page_index)+ " Pages")