import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Function to prompt the user to enter a message template
def enter_template():
    print("Enter your message template with variable names enclosed in curly braces (e.g., {variable_name}):")
    return input("Message Template: ")

# Function to prompt the user to select a file
def select_file():
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return file_path

# Prompt the user to enter a message template
message_template = enter_template()
if not message_template:
    print("No message template provided. Exiting.")
else:
    # Prompt the user to select an Excel file
    file_path = select_file()
    if not file_path:
        print("No file selected.")
    else:
        # Read the selected Excel file into a DataFrame
        df = pd.read_excel(file_path)

        # Generate messages
        messages = []
        for index, row in df.iterrows():
            message = message_template.format(
                company_name=row['Company Name'],
                recipient_name=row['Recipient Name'],
                your_name=row['Your Name'],
                your_contact_info=row['Your Contact Information'],
                specific_interest=row['Specific Interest'],
                specific_skill=row['Specific Skill'],
            )
            messages.append(message)

        # Write messages to a text file
        with open('messages.txt', 'w') as file:
            for message in messages:
                file.write(message + '\n\n')

        print("Messages generated and saved to 'messages.txt'.")
