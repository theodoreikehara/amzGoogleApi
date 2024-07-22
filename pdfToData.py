import PyPDF2
import re
import csv
import os

def extract_name_and_rating(text):
    """
    Extracts the name and rating from the given text.
    
    Args:
    - text (str): The text to extract data from.
    
    Returns:
    - A formatted string with name and rating.
    """
    
    # Using regex to capture the desired patterns
    # Assuming the ID always has 2 characters followed by 12 alphanumeric characters
    pattern = r'^\d+(?P<name>[A-Za-z\s]+[A-Za-z])[A-Z0-9]{2}[A-Z0-9]{12}(?P<rating>Fantastic|Great|Fair|Poor)'
    match = re.search(pattern, text)
    
    # Format the result
    if match:
        name = match.group('name').strip()
        rating = match.group('rating')
        return f"{name}, {rating}"
    else:
        return None

def write_to_csv():
    """
    Reads the content of a file line by line, extracts the desired information, 
    and writes the results to a CSV file.
    """
    
    with open('pdfTemp.txt', 'r') as file, open('name_rating.csv', 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        
        # Write header to CSV
        writer.writerow(["Name", "Rating"])
        
        for line in file:
            result = extract_name_and_rating(line)
            if result:  # Check if result is not None
                name, rating = result.split(', ')
                writer.writerow([name, rating])

def extract_pdf_data(pdf_path):
    """
    Extracts and writes the data between multiple occurrences of two specific strings from a given PDF.
    
    Args:
    - pdf_path (str): Path to the PDF file.
    """
    
    # Open the PDF file in binary reading mode
    with open(pdf_path, 'rb') as file:
        
        # Initialize PDF reader
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Get the total number of pages in the PDF
        num_pages = len(pdf_reader.pages)
        
        # Initialize a variable to hold all the text from the PDF
        full_text = ""
        
        # Loop through all the pages and concatenate text
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            full_text += page.extract_text()

        # Find the start index of the first occurrence
        start_idx = full_text.find('WC-CCSWC-ADDSB DNRPOD Opps.CC Opps.')
        
        # Open the output file in write mode
        with open('pdfTemp.txt', 'w') as output_file:
            while start_idx != -1:  # Continue as long as the start string is found
                # Find the end index after the current start index
                end_idx = full_text.find('Page', start_idx)
                
                if end_idx != -1:  # If end string is found
                    # Adjust the start index to skip the starting string
                    adjusted_start_idx = start_idx + len('WC-CCSWC-ADDSB DNRPOD Opps.CC Opps.')
                    desired_text = full_text[adjusted_start_idx:end_idx]
                    
                    # Split desired_text by newlines and write non-empty lines to output_file
                    for line in desired_text.split('\n'):
                        if line.strip():  # Check if the line is not empty or just whitespace
                            output_file.write(line + "\n")
                    
                    # Move to next possible occurrence by updating the start index
                    start_idx = full_text.find('WC-CCSWC-ADDSB DNRPOD Opps.CC Opps.', end_idx)
                else:  # If end string is not found after current start, exit loop
                    break

def delete_file(file_path):
    """Deletes a file specified by the given path."""
    try:
        os.remove(file_path)
        print(f"{file_path} has been deleted!")
    except FileNotFoundError:
        print(f"{file_path} not found!")
    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    pdf_path = input("Enter the path to the PDF: ")
    extract_pdf_data(pdf_path)
    write_to_csv()
    # delete_file('pdfTemp.txt')
