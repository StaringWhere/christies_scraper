import requests
from bs4 import BeautifulSoup
import bs4
from openpyxl import Workbook
import json
import re
import time
from requests.exceptions import ProxyError


# ---- Parameters ----
# Filename
name = "sanyu"
# URL of lots information JSON (no matter which page)
url = "https://www.christies.com/api/discoverywebsite/search/lot-infos?keyword=sanyu&page=1&is_past_lots=True&sortby=realized_desc&filterids=%7CCoaArtist%7BSanyu%2B(Chang%2BYu)%252c%2B(1901-1966)%7D%7C&language=en"
# Number of pages
maxPage = 7


# Initiate an excel sheet
workbook = Workbook()
sheet = workbook.active

# Sheet head
sheet["B1"] = "Name of painting"
sheet["C1"] = "Estimate"
sheet["D1"] = "Price realised"
sheet["E1"] = "Place"
sheet["F1"] = "Date"
sheet["G1"] = "Auction"
sheet["H1"] = "Details"

# Iteration of pages
lotIndex = 1
for page in range(1, maxPage):

    # Get lots information JSON and decode
    response = requests.get(re.sub(r'(.*page=)[0-9]+(.*)', r'\g<1>' + str(page) + r'\g<2>', url))
    lotsString = response.text
    lotsJSON = json.loads(lotsString)

    # Iteration of lots
    for lotJSON in lotsJSON["lots"]:

        # Unrecognized event type (other than "Sale" and "OnlineSale")
        if lotJSON["event_type"] != "Sale" and lotJSON["event_type"] != "OnlineSale":
            print(str(lotIndex) + " " + "Unrecognized event type")
            lotIndex += 1
            sheet["A" + str(lotIndex + 1)] = lotIndex
            continue

        # Get lot information HTML
        success = False
        maxTry = 10
        while not success and maxTry > 0:

            maxTry -= 1

            try:
                response = requests.get(lotJSON["url"])
                # resonse status is not 200
                if response.status_code != 200:
                    continue
                # respond successfully
                success = True
            
            # Encounter bad connection
            except ProxyError:
                continue
        
        # If failed to get lot information HTML, skip this one
        if not success:
            print(str(lotIndex) + " " + "Bad Connection")
            lotIndex += 1
            sheet["A" + str(lotIndex + 1)] = lotIndex
            continue

        # Parse lot information HTML
        soup = BeautifulSoup(response.text, "html.parser")
        
        nameOfPainting = lotJSON["title_secondary_txt"]
        estimate = lotJSON["estimate_txt"]
        priceRealised = lotJSON["price_realised_txt"]
        date = lotJSON["end_date"][:10]

        auction = ""
        details = ""

        try:
            # Get auction information
            auctionArray = soup.find(class_ = "chr-heading-l-serif").contents
            for item in auctionArray:
                if (type(item) == bs4.element.NavigableString):
                    auction += item
            
            # Get item detaiis
            if lotJSON["event_type"] == "OnlineSale":
                index = re.search(r'{"id":null,"title":"Details"(.*?)}', response.text).span()
                detailsJSON = json.loads(response.text[index[0]: index[1]])
                detailsSoup = BeautifulSoup(detailsJSON["content"], "html.parser")
                detailsArray = detailsSoup.contents
            else:
                detailsArray = soup.find(class_ = "chr-lot-section__accordion--text").contents
            for detail in detailsArray:
                if (type(detail) == bs4.element.NavigableString):
                    details += detail
                elif (type(detail) == bs4.element.Tag and detail.string != None):
                    details += detail.string

        # Can not find the properties
        except AttributeError:
            print(str(lotIndex) + " " + "Wrong format")

        # Write to excel
        sheet["A" + str(lotIndex + 1)] = lotIndex
        sheet["B" + str(lotIndex + 1)] = nameOfPainting
        sheet["C" + str(lotIndex + 1)] = estimate
        sheet["D" + str(lotIndex + 1)] = priceRealised
        sheet["F" + str(lotIndex + 1)] = date
        sheet["G" + str(lotIndex + 1)] = auction.strip()
        sheet["H" + str(lotIndex + 1)] = details

        # Print
        print(str(lotIndex) + " " + "completed")

        # End of each iteration
        lotIndex += 1

# Save excel
workbook.save(filename = name + str(int(time.time())) + ".xlsx")