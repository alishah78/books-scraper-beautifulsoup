import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font

titles = []
prices = []
stocks = []

for page in range(1, 51):
    url = f"https://books.toscrape.com/catalogue/page-{page}.html"
    try:
        response = requests.get(url)
        response.encoding = "utf-8"
    except requests.exceptions.RequestException:
        print(f"Could not reach page {page}, skipping...")
        continue

    soup = BeautifulSoup(response.text, "html.parser")
    books = soup.find_all("article", class_="product_pod")

    for book in books:
        name = book.h3.a["title"].strip()
        if name.strip() == "":
            print("Skipped a book with an empty title.")
            continue

        price_text = book.find("p", class_="price_color").text.strip()
        price = float(price_text.replace("£", ""))
        stock = book.find("p", class_="instock").text.strip()

        titles.append(name)
        prices.append(price)
        stocks.append(stock)

df = pd.DataFrame({"Title": titles, "Price": prices, "Stock": stocks})
df.drop_duplicates(inplace=True)

with pd.ExcelWriter("books_report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Books")
    worksheet = writer.sheets["Books"]
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

print("Excel file created successfully!")
