# Christie's Scraper

Get the details of lots on christie's and write them into an excel sheet.

## Usage

1. Install dependencies

   ```bash
   pip install -r requirements.txt
   ```

2. find the URL of lots information JSON

   It's a fetch request sent during searching, similar to `https://www.christies.com/api/discoverywebsite/search/lot-infos?keyword=...`

3. Modify the parameters in `christies_scraper.py`
4. run `christies_scraper.py`