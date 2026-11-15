import re
from playwright.sync_api import Playwright, sync_playwright
import pandas as pd
import os
import urllib.request
from urllib.parse import quote, urlsplit, urlunsplit
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage

MAX_PAGES = 5


def encode_url_for_request(url: str) -> str:
    parts = urlsplit(url)
    safe_path = "/:@!$&'()*+,;=%-._~"
    safe_query = "/?:@!$&'()*+,;=%-._~"
    return urlunsplit((
        parts.scheme,
        parts.netloc,
        quote(parts.path, safe=safe_path),
        quote(parts.query, safe=safe_query),
        quote(parts.fragment, safe=safe_query),
    ))


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    
    # Block styles and fonts to reduce page load overhead while keeping image URLs available.
    def block_aggressively(route):
        if route.request.resource_type in ["stylesheet", "font"]:
            route.abort()
        else:
            route.continue_()
            
    page.route("**/*", block_aggressively)
    
    print("Navigating to Metalshop...")
    page.goto("https://www.metalshop.cz/tricka-tilka/", wait_until="domcontentloaded")

    consent_button = page.get_by_role("button", name="Souhlasím")
    if consent_button.count() > 0 and consent_button.is_visible():
        consent_button.click()
    
    image_folder = "shirt_images"
    os.makedirs(image_folder, exist_ok=True)

    data_list = []
    page_number = 1
        
    while True:
        print(f"Processing page {page_number}...")

        product_names = page.locator(".product-item-name")
        product_names.first.wait_for(state="visible", timeout=15000)
        product_names_elements = product_names.all()
        
        for name_elem in product_names_elements:
            raw_product_name = name_elem.inner_text().strip()
            
            if not raw_product_name or len(raw_product_name) < 4:
                continue
                
            product_card = name_elem.locator("xpath=../..")
            price_element = product_card.locator("[class*='price'], [class*='cena'], .product-item-price, .price").first
            
            price_text = price_element.inner_text().strip().replace("Kč", "CZK") if price_element.count() > 0 else "Unavailable"
            source_status = product_card.inner_text().lower()
            
            # Prepare a safe filename before checking for an existing image.
            clean_filename = raw_product_name.replace("\n", " ").replace("\r", " ").strip()
            safe_filename = re.sub(r'[\\/*?:"<>|]', "", clean_filename).replace(" ", "_")
            image_filename = f"{safe_filename}.jpg"
            image_path = os.path.join(image_folder, image_filename)
            
            if not (os.path.exists(image_path) and os.path.getsize(image_path) > 0):
                image_url = "No image"
                main_card = name_elem.locator("xpath=../../..")
                image_element = main_card.locator("img").first
                
                if image_element.count() > 0:
                    try:
                        image_element.scroll_into_view_if_needed()
                        img_url_raw = image_element.get_attribute("data-src") or image_element.get_attribute("src")
                        if not img_url_raw or "base64" in img_url_raw.lower() or "placeholder" in img_url_raw.lower():
                            srcset = image_element.get_attribute("data-srcset") or image_element.get_attribute("srcset")
                            if srcset:
                                img_url_raw = srcset.split(",")[-1].strip().split(" ")[0]
                        
                        if img_url_raw:
                            if img_url_raw.startswith("//"):
                                image_url = "https:" + img_url_raw
                            elif img_url_raw.startswith("/"):
                                image_url = "https://metalshop.cz" + img_url_raw
                            else:
                                image_url = img_url_raw
                    except Exception as error:
                        print(f"Could not read image URL for {raw_product_name!r}: {error}")

                if image_url and image_url != "No image" and "base64" not in image_url:
                    try:
                        request = urllib.request.Request(
                            encode_url_for_request(image_url),
                            headers={'User-Agent': 'Mozilla/5.0'}
                        )
                        with urllib.request.urlopen(request, timeout=20) as response, open(image_path, 'wb') as out_file:
                            out_file.write(response.read())
                    except Exception as error:
                        image_filename = "Download failed"
                        print(f"Could not download image for {raw_product_name!r}: {error}")
                else:
                    image_filename = "N/A"

            data_list.append({
                "Product Name": raw_product_name,
                "Price with VAT": price_text,
                "Front image": image_filename,
                "Source status": source_status
            })
            
        next_button = page.locator(".pagination a").filter(
            has_text=re.compile(rf"^\s*{page_number + 1}\s*$")
        ).first
        if next_button.count() > 0 and next_button.is_visible() and page_number < MAX_PAGES:
            print("Next page found; opening it...")
            next_button.click()
            page_number += 1
        else:
            print("Finished collecting products and downloading images.")
            break
            
    # --- DATA CLEANUP & EXCEL EXPORT ---
    if data_list:
        df = pd.DataFrame(data_list)
        df = df.drop_duplicates(subset=["Product Name"]).copy()
        
        def normalize_unavailable(row):
            price = str(row["Price with VAT"]).lower()
            source_status = str(row["Source status"]).lower()
            if any(term in price or term in source_status for term in ("nedostupné", "vyprodáno", "unavailable")):
                return "Unavailable"
            return row["Price with VAT"]
        
        df["Price with VAT"] = df.apply(normalize_unavailable, axis=1)

        def extract_state_and_clean(row):
            product_name = str(row["Product Name"])
            price = str(row["Price with VAT"])
            source_status = str(row["Source status"]).lower()
            
            detected_state = "Standard"
            clean_product_name = product_name
            
            source_status_labels = {
                "BLACK FRIDAY": "Black Friday", "Bestseller": "Bestseller",
                "Novinka": "New", "Poškozené": "Damaged", "Sleva": "Discount", "akce": "Action"
            }
            
            if price == "Unavailable":
                detected_state = "Unavailable"
            else:
                for source_label, english_state in source_status_labels.items():
                    if source_label.lower() in product_name.lower() or source_label.lower() in source_status:
                        detected_state = english_state
            
            discount_match = re.search(r"(\d+)\s*(?:Kč|CZK)", clean_product_name)
            extracted_discount = "N/A"
            if discount_match:
                extracted_discount = f"{discount_match.group(1)} CZK"
                detected_state = "Discount"
                
            if "poškozené" in clean_product_name.lower():
                detected_state = "Damaged"
                
            if detected_state == "Standard" and ("skladem" in source_status or "in stock" in source_status):
                detected_state = "In stock"
            
            clean_product_name = clean_product_name.replace("\n", " ").replace("\r", " ").strip()
            return pd.Series([clean_product_name, detected_state, extracted_discount])

        df[["Product Name", "State", "Discount"]] = df.apply(extract_state_and_clean, axis=1)
        
        def resolve_prices(row):
            base_price = str(row["Price with VAT"]).split("\n")[0].strip()
            discount = str(row["Discount"]).strip()
            if row["State"] == "Unavailable":
                return pd.Series(["Unavailable", "N/A"])
            if row["State"] in ["Discount", "Action"]:
                return pd.Series([base_price, discount])
            return pd.Series([base_price, "N/A"])

        df[["Price with VAT", "Price after discount"]] = df.apply(resolve_prices, axis=1)
        df = df.drop(columns=['Discount', 'Source status'])
        
        state_priority = {
            "Bestseller": 1, "Discount": 2, "Action": 2, "Standard": 3,
            "Black Friday": 3, "New": 3, "In stock": 3, "Damaged": 4, "Unavailable": 5
        }
        df['priority_num'] = df['State'].map(state_priority).fillna(3)
        df = df.sort_values(by=['priority_num', 'Product Name'], ascending=[True, True])
        df = df.drop(columns=['priority_num'])
        
        df = df[["Product Name", "State", "Price with VAT", "Price after discount", "Front image"]]
        
        excel_file = "results_Tshirts.xlsx"
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Shirts")
            worksheet = writer.sheets["Shirts"]
            
            thick_side = Side(border_style="thick", color="000000")
            cell_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)
            center_alignment = Alignment(horizontal="center", vertical="center")
            
            # Use subtle fills to distinguish product states.
            header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
            header_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            product_name_font = Font(name="Calibri", size=14, bold=True)
            state_font = Font(name="Calibri", size=14, underline="single")
            body_font = Font(name="Calibri", size=14)
            left_alignment = Alignment(horizontal="left", vertical="center")
            no_border = Border()
            
            state_fills = {
                "Bestseller": PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid"),
                "Discount": PatternFill(start_color="E5FFCC", end_color="E5FFCC", fill_type="solid"),
                "Action": PatternFill(start_color="E5FFCC", end_color="E5FFCC", fill_type="solid"),
                "New": PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid"),
                "In stock": PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),
                "Standard": PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
                "Damaged": PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid"),
                "Unavailable": PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            }
            
            # Apply formatting and embed images.
            for col in worksheet.columns:
                max_length = 0
                column_letter = col[0].column_letter
                
                for cell in col:
                    # Keep the image column border-free.
                    if column_letter != "E":
                        cell.border = cell_border
                    else:
                        cell.border = no_border
                    
                    if cell.row > 1:
                        worksheet.row_dimensions[cell.row].height = 110
                    else:
                        worksheet.row_dimensions[cell.row].height = 28
                    
                    if cell.row == 1:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = center_alignment
                    else:
                        if column_letter == "A":
                            cell.font = product_name_font
                        elif column_letter == "B":
                            cell.font = state_font
                        else:
                            cell.font = body_font

                        if column_letter in ["B", "C", "D", "E"]:
                            cell.alignment = center_alignment
                        else:
                            cell.alignment = left_alignment
                            
                        current_state = str(worksheet[f"B{cell.row}"].value).strip()
                        
                        if column_letter != "E" and current_state in state_fills:
                            cell.fill = state_fills[current_state]
                    
                    if column_letter == "E" and cell.row > 1 and cell.value and cell.value not in ["N/A", "Download failed"]:
                        image_path = os.path.join(image_folder, str(cell.value))
                        if os.path.exists(image_path):
                            try:
                                img = OpenpyxlImage(image_path)
                                img.height = 140
                                img.width = 100
                                cell.value = "" 
                                worksheet.add_image(img, f"E{cell.row}")
                            except Exception as e:
                                print(f"Could not insert image into Excel: {e}")
                    
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                
                if column_letter == "E":
                    worksheet.column_dimensions[column_letter].width = 16
                else:
                    max_width = 60 if column_letter == "A" else 32
                    worksheet.column_dimensions[column_letter].width = min(max(max_length + 4, 14), max_width)
                    if column_letter == "A":
                        for cell in col:
                            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)



        print("\nDone. The Excel file and image folder are ready.")
    else:
        print("\nNo products were found.")

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
