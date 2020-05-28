# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import re
import urllib

# The years that we are interested in
years = [2011, 2012, 2013, 2014, 2015]
for year in years:
    # for each year we visit the specific url where the excel is located
    url = "https://www.statistics.gr/el/statistics/-/publication/STO04/%s-Q4" % (
        year)
    req = requests.get(url)
    soup = BeautifulSoup(req.text, 'html.parser')
    # we check all tags <a href="" /> in the page's elements and we store them
    rows = soup.findAll('a')  # Extract and return first occurrence of tr
    for row in rows:
        hr = row.getText()
        # for every tag we found we check if the name is the one below, because that the excel sheets we are interested in
        if ("Αφίξεις μη κατοίκων από το εξωτερικό ανά χώρα προέλευσης και μέσο μεταφοράς" in hr):
            # if it matches, then store the url from href and download the excel file from this specific url
            file_url = row.get('href')
            resp = requests.get(file_url)
            # name the excel year.xls where year the year we searched
            output = open('y%s.xls' % (year), 'wb')
            output.write(resp.content)
            output.close()
