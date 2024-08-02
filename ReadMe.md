EtsyAPI_CD3D_01
Only intended for Internal use at Cascadia Digital / Cascadia Design 


Script functions:

Loads the product workbook and worksheet.
Iterates through subdirectories in the product directory, identifying those with valid naming conventions.
Extracts serial numbers (SN) from subdirectory names.
Searches for SN in the Excel sheet.
If found and not already processed, proceeds with creating a new listing on Etsy.
Prepares listing data and sends a request to Etsy API.
Handles the response and extracts the listing ID and web link if successful.
Uploads digital files, images, and videos from designated folders to the listing.
Saves the listing as a draft.
Marks the status in the Excel sheet and saves the workbook.