import os
import pandas as pd

# Define your dictionary for replacement
word_dict = {

}

# Function to replace words in a given text based on the dictionary
def replace_words(text, word_dict):
    for key, value in word_dict.items():
        text = text.replace(key, value)
    return text

# Function to process each Excel file in the folder
def process_excel_files(input_folder, output_folder):
    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, filename)
            df = pd.read_excel(file_path)
            # Replace words in all columns
            df = df.applymap(lambda x: replace_words(str(x), word_dict))
            df.to_excel(output_path, index=False)

# Get the directory of the script
script_dir = os.path.dirname(os.path.realpath(__file__))

# Define input and output folder paths
input_folder = os.path.join(script_dir, "input")
output_folder = os.path.join(script_dir, "output")

# Call the function to process Excel files
process_excel_files(input_folder, output_folder)
