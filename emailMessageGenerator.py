import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Function to prompt the user to enter a custom message template
def enter_template():
    print("Enter your custom message template with variable names enclosed in curly braces (e.g., {variable_name}):")
    return input("Custom Message Template: ")

# Function to use a predefined message template
def use_predefined_template():
    return (
        "Dear {recipient_name},\n\n"
        "My name is {your_name} and I am reaching out to express my interest in {specific_interest} at {company_name}. "
        "I have developed a strong skill set in {specific_skill} and believe I could contribute significantly to your team.\n\n"
        "Please find my contact information below:\n"
        "{your_contact_info}\n\n"
        "Thank you for your time and consideration.\n\n"
        "Best regards,\n"
        "{your_name}"
    )

# Function to prompt the user to select a file
def select_file():
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return file_path

# Prompt the user to choose a message template option
print("Choose a message template option:")
print("1. Enter your own custom message template")
print("2. Use a predefined message template")
option = input("Enter option (1 or 2): ")

if option == '1':
    message_template = enter_template()
    if not message_template:
        print("No message template provided. Exiting.")
    else:
        file_path = select_file()
        if not file_path:
            print("No file selected.")
        else:
            df = pd.read_excel(file_path)
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
            with open('messages.txt', 'w') as file:
                for message in messages:
                    file.write(message + '\n\n')
            print("Messages generated and saved to 'messages.txt'.")

elif option == '2':
    message_template = use_predefined_template()
    file_path = select_file()
    if not file_path:
        print("No file selected.")
    else:
        df = pd.read_excel(file_path)
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
        with open('messages.txt', 'w') as file:
            for message in messages:
                file.write(message + '\n\n')
        print("Messages generated and saved to 'messages.txt'.")
else:
    print("Invalid option. Please choose either 1 or 2.")
