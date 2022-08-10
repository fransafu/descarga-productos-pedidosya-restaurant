import requests
from requests.structures import CaseInsensitiveDict
import pandas as pd
from io import BytesIO
from PIL import Image 
import openpyxl

def get_headers(token=None):
    headers = CaseInsensitiveDict()
    headers["Accept"] = "application/json"
    headers["User-Agent"] = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"

    if token:
        headers["Authorization"] = f"Bearer {token}"

    return headers

def get_token(email, password):
    LOGIN_URL = "https://dlbjm8jos6.execute-api.us-east-1.amazonaws.com/prd/v3/master/login"
    PAYLOAD = { "email": email, "password": password }

    resp = requests.post(LOGIN_URL, json=PAYLOAD, headers=get_headers())
    if resp.status_code != 200:
        return {}

    return resp.json()["accessToken"]

def download_categories(token):
    CATGORIES_URL = f"https://management-api.pedidosya.com/v1/vendors/portal/items/vendors/361790/sections"

    resp = requests.get(CATGORIES_URL, headers=get_headers(token))
    if resp.status_code != 200:
        return []

    return resp.json()

def download_product_list(token, category_id):
    resp = requests.get(f"https://management-api.pedidosya.com/v1/vendors/portal/items/vendors/361790/sections/{category_id}/products", headers=get_headers(token))
    if resp.status_code != 200:
        return []
    return resp.json()

def json_to_excel(products, excel_filename="Products.xlsx"):
    tmp_products = map(lambda product: {
        "category_id": product["category_id"],
        "category_name": product["category_name"],
        "product_id": product["product_id"],
        "product_img": "",
        "product_name": product["product_name"],
        "product_price": product["product_price"],
    }, products)
    df = pd.DataFrame(tmp_products)
    df.to_excel(excel_filename, sheet_name='products', index=False)

    wb = openpyxl.load_workbook(excel_filename)
    ws = wb.worksheets[0]
    index = 2
    ws.column_dimensions['D'].width = 30
    for product in products:
        response = requests.get(product["product_download_img"])

        try:
            im = Image.open(BytesIO(response.content))
            im = im.resize((200,200),Image.NEAREST)
            im.save(product["product_img"])
        except:
            im = Image.open(BytesIO(response.content))
            im = im.resize((200,200),Image.NEAREST)
            im = im.convert('RGB')
            im.save(product["product_img"])

        img = openpyxl.drawing.image.Image(product["product_img"])
        ws.row_dimensions[index].height = 200
        img.anchor = f'D{index}'
        ws.add_image(img)

        index += 1

    wb.save(excel_filename)

def main():
    EMAIL = ""
    PASSWORD = ""

    token = get_token(EMAIL, PASSWORD)
    categories = download_categories(token)

    total_products = []

    for category in categories["sections"]:
        products = download_product_list(token, category["id"])
        for product in products["products"]:
            total_product = {
                "category_id": category["id"],
                "category_name": category["name"],
                "product_id": product["id"],
                "product_img": product["image"]["url"],
                "product_download_img": f'https://images.deliveryhero.io/image/pedidosya/products/{product["image"]["url"]}',
                "product_name": product["name"],
                "product_price": product["price"],
                "product_section_id": product["sectionID"]
            }
            total_products.append(total_product)

    json_to_excel(total_products)

if __name__ == "__main__":
    main()
