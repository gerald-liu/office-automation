import os

# Get the current directory
dir_path = os.getcwd()
# dir_path = 'C:/Users/LWY/Downloads/ilovepdf_split-range'

# Create a new text file in the directory
output_file = open(os.path.join(dir_path, 'output.txt'), 'w')

feedback = '无异议'

i = 1

# Loop through all files in the directory
for filename in os.listdir(dir_path):
    # Check if the file has a .pdf extension
    if filename.endswith('.pdf'):
        # Extract the string between the last "-" and the ".pdf" extension
        start_index = filename.rfind('-') + 1
        if start_index == 0:
            continue
        end_index = filename.rfind('.pdf')
        extracted_string = filename[start_index:end_index]
        # Write the extracted string to the text file
        output_file.write(f'{i}. {extracted_string}: {feedback}\n')
        i += 1

# Close the text file
output_file.close()
