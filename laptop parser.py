import requests
from bs4 import BeautifulSoup
import openpyxl
from tqdm import tqdm
import sys


def main():
    try:
        # Initialize Excel workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Laptops"

        # Set headers
        headers = ["Name", "Price", "Description", "Rating", "Reviews"]
        ws.append(headers)

        base_url = "https://webscraper.io/test-sites/e-commerce/static/computers/laptops"
        current_page = 1
        product_count = 0

        with tqdm(desc="Processing pages") as pbar:
            while True:
                try:
                    # Build URL for current page
                    if current_page == 1:
                        url = base_url
                    else:
                        url = f"{base_url}?page={current_page}"

                    # Fetch page content with timeout
                    response = requests.get(url, timeout=10)
                    response.raise_for_status()  # Raises HTTPError for bad responses

                    soup = BeautifulSoup(response.text, 'html.parser')

                    # Check if page has products
                    products = soup.find_all('div', class_='thumbnail')
                    if not products:
                        break  # No more products found

                    # Process each product
                    for product in products:
                        try:
                            # Extract product details
                            name = product.find('a', class_='title').text.strip()
                            price = product.find('h4', class_='price').text.strip()
                            description = product.find('p', class_='description').text.strip()

                            # Handle rating (count stars)
                            rating = len(product.find_all('span', class_='glyphicon-star'))

                            # Handle reviews count
                            reviews_text = product.find('p', class_='review-count').text.strip()
                            reviews = int(reviews_text.split()[0])  # Extract number from text

                            # Add to Excel
                            ws.append([name, price, description, rating, reviews])
                            product_count += 1

                        except Exception as e:
                            print(f"\nError processing product: {e}", file=sys.stderr)
                            continue

                    current_page += 1
                    pbar.update(1)

                except requests.RequestException as e:
                    print(f"\nError fetching page {current_page}: {e}", file=sys.stderr)
                    break

        # Adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Save Excel file
        wb.save("laptops.xlsx")
        print(f"\nSuccessfully scraped {product_count} products from {current_page - 1} pages.")
        print("Data saved to laptops.xlsx")

    except Exception as e:
        print(f"\nFatal error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()