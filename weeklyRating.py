import csv

# Define the mapping between the rating names and their numerical values
rating_values = {
    "Fantastic": 4,
    "Great": 3,
    "Fair": 2,
    "Poor": 1
}

# Create an inverse mapping from numerical values to rating names
inverse_rating_values = {v: k for k, v in rating_values.items()}

# Initialize a dictionary to store the names, sum of ratings, and count of occurrences
name_ratings = {}

# Read the CSV files and update the dictionary
# Set the base paths for download and desktop directories
msco_path = r'C:\Users\tedik\OneDrive\Desktop\DA Daily Performance MSCO\\'
amz_path = r'C:\Users\tedik\OneDrive\Desktop\DA Daily Performance AMZ\\'

week = '45'
file_name_msco = 'week' + week +'MSCO.csv'
file_name_amz = 'week' + week + 'AMZ.csv'

MSCORating = msco_path + file_name_msco
AMZRating = amz_path + file_name_amz

for file_name in [MSCORating, AMZRating]:
    with open(file_name, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip the header row
        for row in reader:
            if not row:
                continue  # Skip empty rows
            # strip() function used here to remove leading and trailing spaces
            name, rating = [item.strip() for item in row]  
            name_lower = name.lower()  # Convert name to lowercase for case-insensitive comparison
            rating_value = rating_values.get(rating)  # Using get to avoid KeyError if the rating is not found
            if rating_value is None:  # Check if the rating was found in the mapping
                continue  # Skip this row if the rating was not found
            if name_lower in name_ratings:
                name_ratings[name_lower][0] += rating_value
                name_ratings[name_lower][1] += 1
            else:
                name_ratings[name_lower] = [rating_value, 1, name]  # Store the original case name as well

# Calculate the average rating for each name, truncate the decimal part,
# and convert it back to the rating name
for name_lower, (rating_sum, count, original_name) in name_ratings.items():
    average_rating = rating_sum // count
    name_ratings[name_lower] = (inverse_rating_values[average_rating], original_name)

# Create a new CSV file with names and their average ratings
with open("final_rating.csv", mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Name", "Rating"])
    for name_lower, (average_rating, original_name) in name_ratings.items():
        writer.writerow([original_name, average_rating])
