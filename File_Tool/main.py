from frontend import create_window, clear_input
from backend import save_data_to_excel, download_data
import PySimpleGUI as sg
#Importing the functions from the files 

#main function
def main():
    window = create_window()
    submitted_data=[]

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':   #when the user clicks on exit or closes the window 
            break

        if event == 'Submit':    #Submit event
            name = values['name']
            email = values['email']
            phone = values['phone']

            branches = [branch for branch in ['Engineering', 'Medical', 'Degree'] if values[branch]]

            if name and email and phone and branches:
                # If all fields are filled, save the data to Excel
                branch_str = ', '.join(branches)
                save_data_to_excel(name, email, phone, branch_str)  # Save the data to the Excel file
                sg.popup('Data has been saved successfully!')
                submitted_data = [{
                    "Name": name,
                    "Email": email,
                    "Phone": phone,
                    "Branch": branch_str
                }]
            else:
                # Error message will be displayed if there is any missing field.
                sg.popup('All fields are required!', title='Error')

        if event == 'Clear':
            clear_input(window, values)   #This event clears all the values entered on the screen

        if event == 'Download':
            # Call the download function 
            file_format = values['file_format']
            download_data(file_format)

    window.close()  # Close the GUI window when done

if __name__ == "__main__":
    main()  # Run the main function