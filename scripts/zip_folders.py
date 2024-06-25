import os
import py7zr

# Define the directory path
dir_path = 'G:/_Private/EQUITY_DERIVATIVES/Intern/LWY/Archive/20231122'

# Loop through each sub-folder in the directory
for folder_name in os.listdir(dir_path):
    folder_path = os.path.join(dir_path, folder_name)

    # Check if the path is a directory
    if os.path.isdir(folder_path):

        # Define the output file path
        output_path = os.path.join(dir_path, f'{folder_name}.zip')

        # Create a new ZIP archive for the folder
        with py7zr.SevenZipFile(output_path, 'w') as archive:

            # Add all the files in the folder to the archive
            archive.writeall(folder_path, folder_name)
