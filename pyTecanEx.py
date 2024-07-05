## 240610 -updated code to version 3 adding x and std of percent activity and percent inhibition.
## 240608 -updated code to version 2 correcting baseline(t1) and calculate both %active and %inhibition from the original code.
## 240528 -lunch the code as a beta version

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import xlsxwriter
import datetime

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("pyTecanEx: python app for Tecan's data extraction and calculation (version 1.0)")
        self.accepted = False

        # Display license agreement
        self.show_license_agreement()

    def show_license_agreement(self):
        license_text = (
            "Software License Agreement:\n\n"
            "Creative Commons Attribution-NonCommercial-NoDerivs 4.0 International (CC BY-NC-ND 4.0)\n\n"
            "This Software License Agreement (\"Agreement\") is between the software developer (\"Licensor\") and any user (\"Licensee\").\n\n"
            "1. Definitions\n"
            "   - \"Software\": The software product and related materials provided by the Licensor.\n"
            "   - \"Licensor\": The owner of the intellectual property rights to the Software.\n"
            "   - \"Licensee\": Any individual or entity that uses the Software.\n\n"
            "2. License Grant\n"
            "   Subject to this Agreement, the Licensor grants the Licensee a worldwide, royalty-free, non-exclusive, non-transferable license to use the Software, subject to:\n"
            "   a. Attribution: Licensee must credit the Licensor, link to the license, and indicate if changes were made, without implying endorsement.\n"
            "   b. NonCommercial: Licensee may not use the Software for commercial purposes.\n"
            "   c. NoDerivatives: Licensee may not distribute modified versions of the Software.\n\n"
            "3. Termination\n"
            "   This Agreement is effective until terminated. The Licensee's rights terminate automatically if they fail to comply with any terms. Upon termination, all use and distribution must cease.\n\n"
            "4. Disclaimer of Warranties\n"
            "   The Software is provided \"as is\", without warranty of any kind. The Licensor is not liable for any claims or damages arising from its use.\n\n"
            "5. Limitation of Liability\n"
            "   The Licensor is not liable for any special, incidental, indirect, or consequential damages arising from the use or inability to use the Software.\n\n"
            "6. Governing Law\n"
            "   This Agreement is governed by the laws of the Licensor's jurisdiction.\n\n"
            "7. Miscellaneous\n"
            "   If any provision is unenforceable, the remaining provisions remain in effect. This Agreement constitutes the entire agreement regarding the Software.\n\n"
            "8. License Details\n"
            "   This work is licensed under a Creative Commons Attribution-NonCommercial-NoDerivs 4.0 International License. "
            "For the full license, visit: http://creativecommons.org/licenses/by-nc-nd/4.0/\n\n"
            "By using the Software, the Licensee agrees to the terms of this Agreement."
        )

        def on_accept():
            self.accepted = True
            root.deiconify()  # Show the root window after accepting the license
            license_window.destroy()

        license_window = tk.Toplevel(self.root)
        license_window.title("Software License Agreement")
        license_label = tk.Label(license_window, text=license_text, justify="left")
        license_label.pack()
        accept_button = tk.Button(license_window, text="Accept", command=on_accept)
        accept_button.pack()
        root.withdraw()  # Hide the root window until the license is accepted
        self.root.wait_window(license_window)

        # Proceed with application initialization after the license is accepted
        self.initialize_app()

    def initialize_app(self):
        # Set window size
        self.root.geometry("600x400")  # Width x Height

        # Create a label and entry box for cycle variable name
        self.cycle_var_label = tk.Label(self.root, text="Enter cycle number:")
        self.cycle_var_label.pack()
        self.cycle_var_entry = tk.Entry(self.root)
        self.cycle_var_entry.pack()
        
        # Create a label and entry box for column variable name
        self.column_var_label = tk.Label(self.root, text="Enter column number:")
        self.column_var_label.pack()
        self.column_var_entry = tk.Entry(self.root)
        self.column_var_entry.pack()
        
        # Create a label for instructions
        self.label = tk.Label(self.root, text="Please select a file to upload:")
        self.label.pack()

        # Create a button to trigger file upload
        self.upload_button = tk.Button(self.root, text="Upload File", command=self.upload_file)
        self.upload_button.pack()

        # Create an "Excuse" button
        self.excuse_button = tk.Button(self.root, text="Execute", command=self.excuse_print)
        self.excuse_button.pack()

        # Create a "Save" button
        self.save_button = tk.Button(self.root, text="Save", command=self.save_to_excel)
        self.save_button.pack()

        # Initialize variables
        self.cycle_variable = ""
        self.column_variable = ""
        self.df = None
        self.cycle_list = []
        self.cycle_matrices = {}
        self.mean_cycles = []
        self.std_cycles = []
        self.basecor_mean_cycles = []
        self.basecor_std_cycles = []
        self.percent_act_means = []
        self.percent_act_stds = []
        self.percent_int_stds = []
        self.basecor_cycles = {}
        self.percent_acts = {}
        self.percent_ints = {}
        self.percent_int_means = []

        # Create a footnote label
        instructions = (
            "Instructions:\n"
            "1. Enter cycle number and column number according to your protocol.\n"
            "2. Upload an Excel file.\n"
            "3. Click 'Excuse' to process the file.\n"
            "4. Click 'Save' to export data to Excel.\n\n"
            "Developed by Dr. Thanet Pitakbut (Version: 1.0 // 05.07.24)"
        )
        self.footnote_label = tk.Label(self.root, text=instructions, wraplength=500, justify="left", fg="grey")
        self.footnote_label.pack(side="bottom", pady=10)


    def upload_file(self):
        # Open a file dialog to select a file
        file_path = filedialog.askopenfilename()

        if file_path:
            # Read the selected file into a DataFrame using Pandas
            try:
                self.df = pd.read_excel(file_path)  # Change to pd.read_excel() for Excel files
            except Exception as e:
                print("Error:", e)
                # Notify user if there's an error in reading the file

    def excuse_print(self):
        try:
            # Get cycle variable name from entry box
            self.cycle_variable = self.cycle_var_entry.get()

            # Get the maximum number of columns in the DataFrame
            max_columns = self.df.shape[1]

            # Create the cycle list containing the specified number of items
            self.cycle_list = list(range(1, min(int(self.cycle_variable) + 1, max_columns + 1)))

            # Get column variable name from entry box
            self.column_variable = self.column_var_entry.get()
            
            # Display the DataFrame
            print("Data Frame:")
            print(self.df.head())  # Display first few rows of the DataFrame

            # Print the cycle variable name
            print("Cycle variable:", self.cycle_variable)

            # Print the cycle list
            print("Cycle list:", self.cycle_list)

            # Print the column number
            print("Column number:", self.column_variable)

            # Extract data and create matrices
            for i in self.cycle_list:
                # Extract data from the DataFrame and create a list
                cycle_list = list(self.df.iloc[:, i-1])  # Adjusted to zero-based index
                # Print the contents of each cycle list
                print(f"Cycle list {i}: {cycle_list}")

                # Convert the list to a matrix
                self.cycle_matrices[i] = self.list_to_matrix(cycle_list, int(self.column_variable))

                # Print the matrix
                print(f"Cycle matrix {i}:")
                print(self.cycle_matrices[i])
            
            # Calculate mean and standard deviation for each cycle matrix
            self.mean_cycles = []
            self.std_cycles = []
            for k in self.cycle_list:
                if k in self.cycle_matrices:
                    # Calculate mean
                    cycle_mean = np.mean(self.cycle_matrices[k], axis=0)
                    self.mean_cycles.append(cycle_mean.tolist())  # Convert numpy array to list
                    # Calculate standard deviation
                    cycle_std = np.std(self.cycle_matrices[k], axis=0)
                    self.std_cycles.append(cycle_std.tolist())  # Convert numpy array to list

            # Print mean and standard deviation arrays
            print("Mean cycles:", self.mean_cycles)
            print("Standard deviation cycles:", self.std_cycles)
            
            # Additional calculations as provided
            for m in self.cycle_list:
                self.basecor_cycles[m] = self.cycle_matrices[m] - self.cycle_matrices[1]
                print(self.basecor_cycles[m])

            for n in self.cycle_list:
                basecor_cycle_mean = np.mean(self.basecor_cycles[n], axis=0)
                self.basecor_mean_cycles.append(basecor_cycle_mean.tolist())

            print(self.basecor_mean_cycles)

            for o in self.cycle_list:
                basecor_cycle_std = np.std(self.basecor_cycles[o], axis=0)
                self.basecor_std_cycles.append(basecor_cycle_std.tolist())

            print(self.basecor_std_cycles)

            def sublist_last(array):
                last_value = array[:, -1]
                return last_value[:, np.newaxis]

            cycle2 = self.cycle_list[1:]
            for p in cycle2:
                self.percent_acts[p] = self.basecor_cycles[p] / sublist_last(self.basecor_cycles[p])
                print(self.percent_acts[p])

            for q in cycle2:
                percent_act_mean = np.mean(self.percent_acts[q], axis=0)
                self.percent_act_means.append(percent_act_mean.tolist())

            print(self.percent_act_means)

            for r in cycle2:
                percent_act_std = np.std(self.percent_acts[r], axis=0)
                self.percent_act_stds.append(percent_act_std.tolist())

            print(self.percent_act_stds)

            for r in cycle2:
                self.percent_ints[r] = 1 - self.percent_acts[r]
                print(self.percent_ints[r])

            for s in cycle2:
                percent_int_std = np.std(self.percent_ints[s], axis=0)
                self.percent_int_stds.append(percent_int_std.tolist())

            print(self.percent_int_stds)
            
            # Calculate mean percent inhibition
            for r in cycle2:
                percent_int_mean = np.mean(self.percent_ints[r], axis=0)
                self.percent_int_means.append(percent_int_mean.tolist())

            print(self.percent_int_means)
            
            # Write log to file
            self.write_log()

        except Exception as e:
            print("Error:", e)
            
    def list_to_matrix(self, input_list, num_columns):
        try:
            num_rows = len(input_list) // num_columns
            # Ensure that the number of elements in the input list matches the number of rows and columns
            if len(input_list) % num_columns != 0:
                raise ValueError("Number of elements in the list does not match the specified number of columns.")
            # Reshape the input list into a matrix
            matrix = np.array(input_list).reshape(num_rows, num_columns)
            return matrix

        except Exception as e:
            print("Error:", e)

    def write_log(self):
        try:
            current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            log_filename = f"log_{current_time}.txt"
            with open(log_filename, "w") as log_file:
                # Write log information
                log_file.write("Data Frame:\n")
                log_file.write(str(self.df) + "\n\n")
                log_file.write("Cycle variable: " + str(self.cycle_variable) + "\n\n")
                log_file.write("Cycle list: " + str(self.cycle_list) + "\n\n")
                # Write cycle matrices
                for i, matrix in self.cycle_matrices.items():
                    log_file.write(f"Cycle matrix {i}:\n")
                    log_file.write(str(matrix) + "\n\n")
                # Write mean and standard deviation arrays
                log_file.write("Original Mean cycles: " + str(self.mean_cycles) + "\n\n")
                log_file.write("Original Standard deviation cycles: " + str(self.std_cycles) + "\n\n")
                log_file.write("Base corrected mean cycles: " + str(self.basecor_mean_cycles) + "\n\n")
                log_file.write("Base corrected std cycles: " + str(self.basecor_std_cycles) + "\n\n")
                log_file.write("Percent activation means: " + str(self.percent_act_means) + "\n\n")
                log_file.write("Percent activation stds: " + str(self.percent_act_stds) + "\n\n")
                log_file.write("Percent inhibition stds: " + str(self.percent_int_stds) + "\n\n")
                log_file.write("Percent inhibition means: " + str(self.percent_int_means) + "\n\n")
                log_file.write("Log file saved successfully")
            print("Log file saved successfully:", log_filename)
        
        except Exception as e:
            print("Error writing log:", e)
    
    def save_to_excel(self):
        try:
            # Save mean cycles to a separate Excel file
            if self.mean_cycles:
                workbook_mean = xlsxwriter.Workbook("ori_cycles_mean.xlsx")
                worksheet_mean = workbook_mean.add_worksheet()
                for row_num, mean_array in enumerate(self.mean_cycles):
                    worksheet_mean.write_row(row_num, 0, mean_array)
                workbook_mean.close()
            
            # Save standard deviation cycles to a separate Excel file
            if self.std_cycles:
                workbook_std = xlsxwriter.Workbook("ori_cycles_std.xlsx")
                worksheet_std = workbook_std.add_worksheet()
                for row_num, std_array in enumerate(self.std_cycles):
                    worksheet_std.write_row(row_num, 0, std_array)
                workbook_std.close()

            # Save base corrected mean cycles to a separate Excel file
            if self.basecor_mean_cycles:
                workbook_basecor_mean = xlsxwriter.Workbook("basecor_mean_cycles.xlsx")
                worksheet_basecor_mean = workbook_basecor_mean.add_worksheet()
                for row_num, mean_array in enumerate(self.basecor_mean_cycles):
                    worksheet_basecor_mean.write_row(row_num, 0, mean_array)
                workbook_basecor_mean.close()

            # Save base corrected std cycles to a separate Excel file
            if self.basecor_std_cycles:
                workbook_basecor_std = xlsxwriter.Workbook("basecor_std_cycles.xlsx")
                worksheet_basecor_std = workbook_basecor_std.add_worksheet()
                for row_num, std_array in enumerate(self.basecor_std_cycles):
                    worksheet_basecor_std.write_row(row_num, 0, std_array)
                workbook_basecor_std.close()

            # Save percent activation means to a separate Excel file
            if self.percent_act_means:
                workbook_percent_act_mean = xlsxwriter.Workbook("percent_act_means.xlsx")
                worksheet_percent_act_mean = workbook_percent_act_mean.add_worksheet()
                for row_num, mean_array in enumerate(self.percent_act_means):
                    worksheet_percent_act_mean.write_row(row_num, 0, mean_array)
                workbook_percent_act_mean.close()

            # Save percent activation stds to a separate Excel file
            if self.percent_act_stds:
                workbook_percent_act_std = xlsxwriter.Workbook("percent_act_stds.xlsx")
                worksheet_percent_act_std = workbook_percent_act_std.add_worksheet()
                for row_num, std_array in enumerate(self.percent_act_stds):
                    worksheet_percent_act_std.write_row(row_num, 0, std_array)
                workbook_percent_act_std.close()

            # Save percent inhibition stds to a separate Excel file
            if self.percent_int_stds:
                workbook_percent_int_std = xlsxwriter.Workbook("percent_int_stds.xlsx")
                worksheet_percent_int_std = workbook_percent_int_std.add_worksheet()
                for row_num, std_array in enumerate(self.percent_int_stds):
                    worksheet_percent_int_std.write_row(row_num, 0, std_array)
                workbook_percent_int_std.close()

            # Save percent inhibition means to a separate Excel file
            if self.percent_int_means:
                workbook_percent_int_mean = xlsxwriter.Workbook("percent_int_means.xlsx")
                worksheet_percent_int_mean = workbook_percent_int_mean.add_worksheet()
                for row_num, mean_array in enumerate(self.percent_int_means):
                    worksheet_percent_int_mean.write_row(row_num, 0, mean_array)
                workbook_percent_int_mean.close()

            # Save individual cycle matrices to separate Excel files
            for j in self.cycle_list:
                if j in self.cycle_matrices:
                    # Get the cycle matrix for current cycle
                    cycle_matrix = self.cycle_matrices[j]

                    # Create a new Excel workbook
                    workbook = xlsxwriter.Workbook(f"ori_cycle_{j}.xlsx")
                    worksheet = workbook.add_worksheet()

                    # Write data from the cycle matrix to the worksheet
                    for row_num, row_data in enumerate(cycle_matrix):
                        worksheet.write_row(row_num, 0, row_data.tolist())  # Convert numpy array to list

                    # Write mean and standard deviation data
                    worksheet.write_row(len(cycle_matrix), 0, ["Mean"] + self.mean_cycles[j-1])
                    worksheet.write_row(len(cycle_matrix) + 1, 0, ["Standard Deviation"] + self.std_cycles[j-1])

                    # Close the workbook
                    workbook.close()

            for k in self.cycle_list:
                if k in self.basecor_cycles:
                    # Get the cycle matrix for current cycle
                    basecor_cycle = self.basecor_cycles[k]

                    # Create a new Excel workbook
                    workbook = xlsxwriter.Workbook(f"basecor_cycle_{k}.xlsx")
                    worksheet = workbook.add_worksheet()

                    # Write data from the cycle matrix to the worksheet
                    for row_num, row_data in enumerate(basecor_cycle):
                        worksheet.write_row(row_num, 0, row_data.tolist())  # Convert numpy array to list

                    # Write mean and standard deviation data
                    worksheet.write_row(len(basecor_cycle), 0, ["Mean"] + self.basecor_mean_cycles[k-1])
                    worksheet.write_row(len(basecor_cycle) + 1, 0, ["Standard Deviation"] + self.basecor_std_cycles[k-1])

                    # Close the workbook
                    workbook.close()

            for l in self.cycle_list:
                if l in self.percent_acts:
                    # Get the cycle matrix for current cycle
                    percent_act = self.percent_acts[l]

                    # Create a new Excel workbook
                    workbook = xlsxwriter.Workbook(f"percent_act_{l}.xlsx")
                    worksheet = workbook.add_worksheet()

                    # Write data from the cycle matrix to the worksheet
                    for row_num, row_data in enumerate(percent_act):
                        worksheet.write_row(row_num, 0, row_data.tolist())  # Convert numpy array to list

                    # Write mean and standard deviation data
                    worksheet.write_row(len(percent_act), 0, ["Mean"] + self.percent_act_means[l-2])
                    worksheet.write_row(len(percent_act) + 1, 0, ["Standard Deviation"] + self.percent_act_stds[l-2])

                    # Close the workbook
                    workbook.close()

            for m in self.cycle_list:
                if m in self.percent_ints:
                    # Get the cycle matrix for current cycle
                    percent_int = self.percent_ints[m]

                    # Create a new Excel workbook
                    workbook = xlsxwriter.Workbook(f"percent_int_{m}.xlsx")
                    worksheet = workbook.add_worksheet()

                    # Write data from the cycle matrix to the worksheet
                    for row_num, row_data in enumerate(percent_int):
                        worksheet.write_row(row_num, 0, row_data.tolist())  # Convert numpy array to list

                    # Write mean and standard deviation data
                    worksheet.write_row(len(percent_int), 0, ["Mean"] + self.percent_int_means[m-2])
                    worksheet.write_row(len(percent_int) + 1, 0, ["Standard Deviation"] + self.percent_int_stds[m-2])

                    # Close the workbook
                    workbook.close()

            print("Data saved to Excel files successfully")
        
        except Exception as e:
            print("Error saving to Excel:", e)

# Create the main window
root = tk.Tk()

# Create the application
app = App(root)

# Add protocol to handle window closure event
root.protocol("WM_DELETE_WINDOW", root.destroy)


# Run the main event loop
root.mainloop()

# %%
