# 💻 Laptops Scraper

This project scrapes laptop product data from [webscraper.io test e-commerce site](https://webscraper.io/test-sites/e-commerce/allinone/computers/laptops).

## Features

- Extracts product name, price, description, rating, reviews count
- Follows links to get full product descriptions
- Multi-threaded for performance
- Saves to `laptops_detailed.xlsx`

## Requirements

- Python 3.7+
- `requests`
- `beautifulsoup4`
- `openpyxl`
- `tqdm`
- `concurrent.futures`

## Usage

```bash
pip install -r requirements.txt
python laptop\ parcer.py
```

Output: `laptops_detailed.xlsx`
