# books-ecommerce-scraper
Scrapes 1000 books from books.toscrape.com and delivers structured Excel report
# Book Data Scraper

## What it does
Scrapes 1000 books across 50 pages from books.toscrape.com
and generates a structured Excel report.

## Libraries Used
- Python
- BeautifulSoup
- Requests
- Pandas
- OpenPyXL

## Features
- Scrapes 1000 books across 50 pages (pagination)
- Extracts book name, price, rating, availability
- Cleans and removes duplicate data
- Generates Excel report with 3 sheets:
  - Books Data (all 1000 books)
  - Summary (total books, avg price, rating breakdown)
  - Top Rated Books (4 and 5 star only)

## Output Screenshot
<img width="1515" height="810" alt="image" src="https://github.com/user-attachments/assets/9e36446a-fa14-4957-b68b-7c24c5033375" />
<img width="1514" height="810" alt="image" src="https://github.com/user-attachments/assets/e3f14c06-0b8e-4bf9-a63b-ac6f8b1738c7" />
<img width="1511" height="803" alt="image" src="https://github.com/user-attachments/assets/f01f8335-ebe0-437c-a6d4-88c16fd94183" />

