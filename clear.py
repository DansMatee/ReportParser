import os

# Define paths
main_folder = os.path.dirname(os.path.abspath(__file__))
filtered_output = os.path.join(main_folder, 'filtered_output.csv')
reports_folder = os.path.join(main_folder, 'reports')

os.remove('filtered_output.csv')

# Delete filtered_output file if it exists
if os.path.exists(filtered_output):
    os.remove(filtered_output)
    print(f'Deleted {filtered_output}')
else:
    print(f'{filtered_output} not found.')

# Delete files within each subfolder in Reports
if os.path.exists(reports_folder):
    for folder_name in os.listdir(reports_folder):
        folder_path = os.path.join(reports_folder, folder_name)
        if os.path.isdir(folder_path):  # Check if it's a directory
            for file_name in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file_name)
                if os.path.isfile(file_path):
                    os.remove(file_path)  # Delete each file
                    print(f'Deleted file: {file_path}')
else:
    print(f'{reports_folder} not found.')

print("Cleanup complete.")