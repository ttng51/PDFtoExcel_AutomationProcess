#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import fitz  
import pandas as pd
import re

folder_path = '.....'

# List all PDF files in the folder
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# Define Excel file path
excel_path = '......'

# Create a Pandas Excel writer using the XlsxWriter as the engine.
excel_writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
# Function to combine elements with a comma at the end with the next element
def combine_elements_with_comma(data_list):
    combined_list = []
    skip_next = False

    for i, item in enumerate(data_list):
        # If the previous flag is true, skip this iteration
        if skip_next:
            skip_next = False
            continue
        # Check if the current item has a comma at the end
        if item.endswith(',') and i + 1 < len(data_list):
            # Combine with the next element
            combined_list.append(item + ' ' + data_list[i + 1])
            # Set the flag to skip the next element as it has been combined
            skip_next = True
        else:
            combined_list.append(item)
            
    return combined_list
# Function to combine elements with the next element if the next is a non-empty, non-dash string and not a number
def combine_elements_with_next_string(data_list): 
    combined_list = []
    i = 0

    while i < len(data_list) - 1:
        current_item = data_list[i].strip()
        next_item = data_list[i + 1].strip()

        # Check if both current and next items contain alphabetic characters
        if any(char.isalpha() for char in current_item) and any(char.isalpha() for char in next_item):
            combined_list.append(f"{current_item} {next_item}")
            i += 2  # Skip the next item as it's combined
        else:
            combined_list.append(current_item)
            i += 1

    # Add the last item if it wasn't combined
    if i == len(data_list) - 1:
        combined_list.append(data_list[i].strip())

    return combined_list
empty_pdf_files = []
for pdf_file in pdf_files:
    # Initialize variables to store the page number where the broad search term is found
    found_page_numbers = []
    data_frames = []  # List to hold data frames of each page
    # Construct the full path to the PDF file
    pdf_path = os.path.join(folder_path, pdf_file)
    
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    # Define the broad search term
    broad_search_terms = [
    'nach Segmenten', 'Modellreihen', 'Kraftstoffarten', 
    'CO2-Emissionen', 'Kraftstoffverbrauch', 'Neuzulassungen von Personenkraftwagen'
    ]

    # Search through the document for the broad search terms
    for page_number in range(len(pdf_document)):
        # Extract text from the current page
        page_text = pdf_document[page_number].get_text()

        # Check if all the broad search terms are in the current page's text
        if all(term in page_text for term in broad_search_terms):
            found_page_numbers.append(page_number)
    # Define the regex pattern for matching the month and year
    month_year_pattern = r'Neuzulassungen von Personenkraftwagen im (\w+) (\d{4})'

    # Initialize the variable to store the found month and year
    found_month_year = None

    # Search through the document for the month and year
    for page_number in range(len(pdf_document)):
        # Extract text from the current page
        page_text = pdf_document[page_number].get_text()

        # Use regex to find the first occurrence of the month and year pattern
        match = re.search(month_year_pattern, page_text)

        # If a match is found, store it and break out of the loop
        if match:
            found_month_year = match.groups()
            break  # Assuming we only want the first match

    # Format the result as 'Month_Year'
    if found_month_year:
        formatted_month_year = f"{found_month_year[0]}_{found_month_year[1]}"
    else:
        formatted_month_year = "No match found"
    # List of page numbers to iterate over
    pages_to_iterate = found_page_numbers

    for page_number in pages_to_iterate:
        page = pdf_document[page_number]

        # Get the page dimensions
        page_rect = page.rect
        page_width = page_rect.width
        page_height = page_rect.height

        # Estimate the height of the non-footnote area (for example, 90% of the page)
        non_footnote_area_height = page_height * 0.95

        # Define the rectangle for the non-footnote area
        clip_rect = fitz.Rect(0, 0, page_width, non_footnote_area_height)

        # Extract text from the defined area
        page_text1 = page.get_text("text", clip=clip_rect)
        
        # Split the text into lines
        lines_page1 = page_text1.strip().split('\n')
        
        # Remove lines that contain the footnote indicator "1)"
        lines_page = [line for line in lines_page1 if "1)" not in line]
        
        # Find the index of '13' and set it to None if not found
        index_of_13 = lines_page.index('13') if '13' in lines_page else None

        # Define the key phrases
        key_phrase = 'Kraftstoffarten, CO2-Emissionen und Kraftstoffverbrauch'
        segment_key = 'Segment/ '

        # Find the index of the key phrases
        index_of_key_phrase = lines_page.index(key_phrase) if key_phrase in lines_page else None
        index_of_segment = lines_page.index(segment_key) if segment_key in lines_page else None

        # Check if '13' is at the end of the list
        if index_of_13 is not None and index_of_13 == len(lines_page) - 1:
            # If the key phrase is found, set the start index after it
            if index_of_key_phrase is not None and index_of_segment is not None:
                # Set the start index after 'Kraftstoffarten, CO2-Emissionen und Kraftstoffverbrauch'
                start_index = index_of_key_phrase + 1
                # Set the end index at 'Segment/ '
                end_index = index_of_segment
                # Extract the lines between the key phrase and 'Segment/ '
                lines_clean = lines_page[start_index:end_index]
            else:
                # If either key phrase is not found, set lines_clean to an empty list
                lines_clean = []
        else:
            # If '13' is not at the end, or '13' is not found, use the original logic for lines_clean
            # Check if the string at the index after '13' is empty
            if index_of_13 is not None and index_of_13 < len(lines_page) - 1 and not lines_page[index_of_13 + 1].strip():
                # If the string is empty, skip it
                start_index = index_of_13 + 2
            elif index_of_13 is not None:
                start_index = index_of_13 + 1
            else:
                start_index = 0

            # Set lines_clean starting from index_of_13 or from the beginning if '13' is not found
            lines_clean = lines_page[start_index:] if index_of_13 is not None else lines_page

        strings_to_delete = [
            'OBERE MITTELKLASSE', 'MITTELKLASSE', 'SUVs', 'MINI-VANS', 'GROSSRAUM-VANS',
            'UTILITIES', 'GELÄNDEWAGEN', 'SPORTWAGEN', 'WOHNMOBILE', 'NEUZULASSUNGEN',
            'OBERKLASSE', 'KOMPAKTKLASSE', 'KLEINWAGEN', 'MINIS']

        try:
            # Find the index of 'NEUZULASSUNGEN'
            index_of_neuzulassungen = lines_clean.index('NEUZULASSUNGEN')

            # Remove the empty string before 'NEUZULASSUNGEN', if it exists
            if index_of_neuzulassungen > 0 and not lines_clean[index_of_neuzulassungen - 1].strip():
                del lines_clean[index_of_neuzulassungen - 1]
                # Adjust index after deletion
                index_of_neuzulassungen -= 1

            # Remove 'NEUZULASSUNGEN' itself
            del lines_clean[index_of_neuzulassungen]

            # Check for the next item after 'NEUZULASSUNGEN' was removed, if it is empty string, remove it
            if not lines_clean[index_of_neuzulassungen].strip():
                del lines_clean[index_of_neuzulassungen]
        except ValueError:
            pass

        # Delete the specific strings from the list
        lines = [item for item in lines_clean if item not in strings_to_delete]
        lines1 = [item for i, item in enumerate(lines) if not (item == ' ' and i < len(lines) - 1 and lines[i + 1] == 'ZUSAMMEN ')]

        lines2 = []
        i = 0
        while i < len(lines1):
            # Check if we are at a potential position for ' INSGESAMT '
            if i + 2 < len(lines1) and lines1[i + 2].strip() == 'INSGESAMT':
                # Check the two preceding items for empty strings
                if lines1[i].strip() == '' and lines1[i + 1].strip() == '':
                    i += 2  # Skip the two empty strings and do not add to the filtered list
                    continue
            lines2.append(lines1[i])
            i += 1

        lines3 = combine_elements_with_comma(lines2)
        lines4 = combine_elements_with_next_string(lines3)

        # Split the data into chunks of `num_columns` size
        num_columns = 14
        chunks = [lines4[i:i + num_columns] for i in range(0, len(lines4), num_columns)]
        # Create the DataFrame with filtered rows and correct column headers
        table_headers = [
        "Segment/Modellreihe", "Insgesamt Anzahl", "Insgesamt CO2-Emission in g/km", "Benzin Anzahl",
        "Benzin CO2-Emission in g/km", "Benzin Kraftstoffverbrauch in l/100km", "Diesel Anzahl",
        "Diesel CO2-Emission in g/km", "Diesel Kraftstoffverbrauch in l/100km", "Erdgas(CNG)(einschl.bivalent)",
        "Flüssiggas(LPG)(einschl.bivalent)", "Hybrid Ingesamt", "Hybrid dar.Plug-in",
        "Hybrid Elektro"]
        try:
            # Append DataFrame of current page to the list
            data_frames.append(pd.DataFrame(chunks, columns=table_headers))
            df = pd.DataFrame(chunks, columns=table_headers)
        except ValueError as e:
            continue
    if df.empty:
        empty_pdf_files.append(pdf_file)
    # Concatenate all DataFrames into one
    try:
        final_df = pd.concat(data_frames, ignore_index=True)
        for column in final_df.columns[1:]:
            final_df[column] = final_df[column].astype(str).str.replace(' ', '', regex=False)
            final_df[column] = final_df[column].astype(str).str.replace(',', '.', regex=False)
        final_df.to_excel(excel_writer, sheet_name=formatted_month_year, index=False)
    except ValueError as e:
        continue
    # Close the PDF document
    pdf_document.close()
# Close the Pandas Excel writer and save the Excel file
excel_writer.save()
#print("PDF files with empty dataframes:", empty_pdf_files)

