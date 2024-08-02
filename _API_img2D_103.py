import os
import openpyxl
import requests
import json
from requests_oauthlib import OAuth1

# Etsy API Configuration
API_KEY = 'g075bclqi7x6rt59dng388vw'
SHARED_SECRET = 'pf2ob9cnde'
oauth = OAuth1(API_KEY, client_secret=SHARED_SECRET)

config = {
    'dir_products': r'C:\Users\Casey\OneDrive\CSM\CD\CD_laser\CDL_products',
    'dir_products_posted': r'C:\Users\Casey\OneDrive\CSM\CD\CD_laser\CDL_products\_posted',
    'sheet_listings': r"C:\Users\Casey\OneDrive\CSM\CD\Documents\ListingTitles_STL.xlsx",
}

template_listing_title = "Fossil Full Pterodactyl 3D STL File for CNC Router, 3D Print, Casting, Wood Carving Engraving 3D Relief Dino dinosaur animal skeleton 1269"

def create_etsy_listing():
    print("Loading Excel workbook...")
    # Load the workbook and worksheet
    try:
        wb = openpyxl.load_workbook(config['sheet_listings'], data_only=True)
        ws = wb.active
        print("Workbook loaded successfully.")
    except Exception as e:
        print(f"Error loading workbook: {e}")
        return

    # Look for subfolders starting with a 4-digit number followed by a space
    print(f"Looking in directory: {config['dir_products']}")
    for subdir in os.listdir(config['dir_products']):
        if os.path.isdir(os.path.join(config['dir_products'], subdir)) and subdir[:4].isdigit() and subdir[4] == ' ':
            subdir_product = os.path.join(config['dir_products'], subdir)
            SN = subdir[:4]
            print(f"Found product subdirectory: {subdir_product}, SN: {SN}")

            # Search for the SN in the Excel sheet
            print(f"Searching for SN {SN} in Excel sheet...")
            found_sn = False
            for row in ws.iter_rows(min_row=2, max_col=12):
                cell_value = str(row[0].value).strip()  # Convert to string and strip any extra spaces
                listing_title = row[3].value
                listing_section = row[5].value
                status = row[11].value

                # print(f"Checking row {row[0].row}: SN={cell_value}, Listing={listing_title}, Section={listing_section}, Status={status}")

                if cell_value == SN:
                    found_sn = True
                    print(f"Found matching SN in sheet: {SN}")
                    if status is None:
                        print(f"Cell in the 11th column is empty, proceeding with listing creation for SN: {SN}")
                        print(f"Listing Title: {listing_title}, Listing Section: {listing_section}")

                        # Create a new listing
                        try:
                            new_listing_data = {
                                'title': listing_title,
                                'description': "Your description here",
                                'price': '20.00',  # Set your price
                                'quantity': 10,    # Set your quantity
                                'shipping_template_id': 123456789,  # Replace with your shipping template ID
                                'shop_section_id': listing_section,  # Use the appropriate section ID
                                'taxonomy_id': 123,  # Replace with the correct taxonomy ID
                                'who_made': 'i_did',
                                'is_supply': 'false',
                                'when_made': '2020_2022',
                                'item_weight': '0.1',
                                'item_length': '10',
                                'item_width': '10',
                                'item_height': '1',
                                'item_weight_unit': 'oz',
                                'item_dimensions_unit': 'in',
                                'state': 'draft'
                            }

                            url = "https://openapi.etsy.com/v2/listings"
                            headers = {"Content-Type": "application/x-www-form-urlencoded"}
                            response = requests.post(url, headers=headers, data=new_listing_data, auth=oauth)
                            response_data = response.json()

                            if response.status_code == 201:
                                listing_id = response_data['results'][0]['listing_id']
                                print(f"Created new listing, ID: {listing_id}")
                            else:
                                print(f"Failed to create listing: {response_data['error']}")
                                continue

                            # Remove all images, videos, and digital file uploads from the listing
                            # Etsy API does not support deleting all media at once, so we skip this part

                            # Upload new files from subdir_product/_upload
                            upload_folder = os.path.join(subdir_product, '_upload')
                            if os.path.exists(upload_folder):
                                print(f"Uploading digital files from: {upload_folder}")
                                for file in os.listdir(upload_folder):
                                    file_path = os.path.join(upload_folder, file)
                                    with open(file_path, 'rb') as f:
                                        files = {'file': f}
                                        response = requests.post(f"https://openapi.etsy.com/v2/listings/{listing_id}/files", files=files, auth=oauth)
                                        print(f"Uploaded digital file: {file_path}, response: {response.status_code}")
                            else:
                                print(f"No upload folder found at: {upload_folder}")

                            # Upload images and videos from subdir_product (no subfolders)
                            print(f"Uploading images and videos from: {subdir_product}")
                            for file in os.listdir(subdir_product):
                                file_path = os.path.join(subdir_product, file)
                                if file_path.endswith(('.jpg', '.png')):
                                    with open(file_path, 'rb') as f:
                                        files = {'image': f}
                                        response = requests.post(f"https://openapi.etsy.com/v2/listings/{listing_id}/images", files=files, auth=oauth)
                                        print(f"Uploaded image: {file_path}, response: {response.status_code}")
                                elif file_path.endswith(('.mp4', '.mov')):
                                    with open(file_path, 'rb') as f:
                                        files = {'video': f}
                                        response = requests.post(f"https://openapi.etsy.com/v2/listings/{listing_id}/videos", files=files, auth=oauth)
                                        print(f"Uploaded video: {file_path}, response: {response.status_code}")

                            # Save the listing as a draft
                            update_data = {'state': 'draft'}
                            response = requests.put(f"https://openapi.etsy.com/v2/listings/{listing_id}", data=update_data, auth=oauth)
                            print(f"Saved new listing as draft, listing ID: {listing_id}")

                            # Mark the 11th column in the sheet with "X"
                            row[10].value = "X"
                            wb.save(config['sheet_listings'])
                            print(f"Marked listing as processed in the Excel sheet for SN: {SN}")
                        except Exception as e:
                            print(f"Error creating or uploading to listing for SN: {SN}. Error: {e}")
                    else:
                        print(f"Cell in the 11th column is not empty, skipping SN: {SN}")
                    break
            
            if not found_sn:
                print(f"SN {SN} not found in the Excel sheet.")

if __name__ == "__main__":
    create_etsy_listing()
