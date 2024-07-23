from frontend import create_window  # Import function to create the GUI window
from backend import create_excel_file, save_data_to_excel, download_data  # Import functions for file operations
import PySimpleGUI as sg  # Import PySimpleGUI for GUI functionality

def main():
    """
    Main function to run the application.
    - Initializes the Excel file.
    - Creates the GUI window.
    - Handles events such as submitting data, clearing inputs, and downloading files.
    """
    create_excel_file()  # Ensure the Excel file exists and is properly set up
    window = create_window()  # Create and display the GUI window
    submitted_data = []  # List to keep track of the submitted data

    while True:
        # Read the event and values from the GUI window
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            # Exit the application if the window is closed or 'Exit' button is pressed
            break

        if event == 'Submit':
            # Handle the 'Submit' button click event
            name = values['name']
            email = values['email']
            phone = values['phone']
            # Get selected branches from the values dictionary
            branches = [branch for branch in ['Engineering', 'Medical', 'Degree'] if values[branch]]

            if name and email and phone and branches:
                # If all fields are filled, save the data to Excel and update the submitted_data list
                branch_str = ', '.join(branches)  # Convert the list of branches to a comma-separated string
                save_data_to_excel(name, email, phone, branch_str)  # Save the data to the Excel file
                sg.popup('Data saved successfully!')  # Show a success popup message
                submitted_data = [{
                    "Name": name,
                    "Email": email,
                    "Phone": phone,
                    "Branch": branch_str
                }]
            else:
                # Show an error popup message if any required field is missing
                sg.popup('All fields are required!', title='Error')

        if event == 'Clear':
            # Handle the 'Clear' button click event
            clear_input(window, values)  # Clear all input fields in the GUI

        if event == 'Download':
            # Handle the 'Download' button click event
            if submitted_data:
                # If there is data to download, call the download function
                file_format = values['file_format']
                download_data(file_format, submitted_data)
            else:
                # Show an error popup message if no data has been entered
                sg.popup('Enter the Data First.', title='Error')

    window.close()  # Close the GUI window when done

if __name__ == "__main__":
    main()  # Run the main function if this script is executed directly