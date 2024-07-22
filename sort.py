import csv

# Define a custom sort order
order = {'Fantastic': 0, 'Great': 1, 'Fair': 2, 'Poor': 3}

def sort_csv(input_filename, output_filename):
    # Read the data from the CSV file
    with open(input_filename, 'r') as f:
        reader = csv.reader(f)
        header = next(reader)  # extract the header
        data = list(reader)

    # Filter and sort the data by rating
    filtered_sorted_list = sorted(
        [row for row in data if row[1] in order],
        key=lambda row: order[row[1]]
    )

    # Write the sorted data to the output CSV file
    with open(output_filename, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)  # write the header
        writer.writerows(filtered_sorted_list)

# Usage
input_file = "final_rating.csv"
output_file = "sort_final_rating.csv"
sort_csv(input_file, output_file)