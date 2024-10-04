from Esmart_ACS import Esmart_ACS
import pandas as pd
import tkinter as tk
from tabulate import tabulate
from tkinter import filedialog, messagebox, Label, Button, Text, Entry, StringVar, OptionMenu
from tkinter import ttk  # Use ttk for styling widgets
import threading
import time  # To simulate the processing time



class CheckingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Order and Packing List Checker 新贵集团 Newmen 内部订单")
        self.root.geometry("1200x800")


        # Make the window resizable
        self.root.grid_rowconfigure(0, weight=0)  # Title row

        self.root.grid_columnconfigure(0, weight=1)
        

        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 12), padding=10)
        style.configure('TLabel', font=('Helvetica', 12), background="#f0f0f0", foreground="blue")
        style.configure('Title.TLabel', font=('Helvetica', 16, 'bold'), foreground="darkred")
        style.configure('TFrame', background="#f0f0f0")

        root.grid_rowconfigure(0, weight=1)  # Allow the row to expand
        root.grid_columnconfigure(0, weight=1)  # Left column (main_frame) expands
        root.grid_columnconfigure(1, weight=2)  # Right column (second_frame) expands more

        self.language_var = StringVar()
        self.language_var.set("en")  # Default to English

        

        # Title Label
        title_label = ttk.Label(root, text="Order and Packing List Checker 新贵集团 Newmen", style='Title.TLabel', anchor='center')
        title_label.grid(row=3, column=0, pady=0, padx=0)  # Use grid instead of pack
        # Add "Back to Launcher" button to go back to startup.py
        self.back_button = Button(root, text="Back to Launcher", command=self.back_to_launcher, font=('Helvetica', 13))
        self.back_button.grid(row=3, column=1, padx=0)
        

        main_frame = tk.Frame(root, padx=50, pady=0,)
        main_frame.grid(row=0, column=0, sticky='new')  # Use grid instead of pack

        # Order and Packing Files Frame
        files_frame = tk.LabelFrame(main_frame, text="File Selection", padx=10, pady=10,font=('Helvetica', 18))
        files_frame.grid(row=0, column=0, sticky='ew', pady=0)

        

        # File selection buttons    
        self.select_order_files_button = Button(files_frame, text="Select Order Files", command=self.select_order_files, font=('Helvetica', 12))
        self.select_order_files_button.grid(row=0, column=0, padx=10, pady=5, sticky='ew')

        self.select_packing_file_button = Button(files_frame, text="Select Packing File", command=self.select_packing_file, font=('Helvetica', 12))
        self.select_packing_file_button.grid(row=0, column=1, padx=10, pady=5, sticky='ew')

        # File info labels
        self.order_files_label = Label(files_frame, text="Order Files: None", font=('Helvetica', 12))
        self.order_files_label.grid(row=1, column=0, columnspan=2, sticky='w', padx=10)

        self.packing_file_label = Label(files_frame, text="Packing File: None" , font=('Helvetica', 12))
        self.packing_file_label.grid(row=2, column=0, columnspan=2, sticky='w', padx=10)

        # Search Frame
        search_frame = tk.LabelFrame(main_frame, text="Search and Filter", padx=10, pady=10, font=('Helvetica', 18))
        search_frame.grid(row=0, column=1, sticky='ew', pady=10, padx=40)

        # Search entry and button
        self.search_entry = Entry(search_frame, width=40, font=('Helvetica', 12))
        self.search_entry.insert(0, 'Type to search...')  # Insert placeholder text
        self.search_entry.grid(row=0, column=0, padx=10, pady=5, sticky='w')

        self.search_button = Button(search_frame, text="Search", command=self.search_product_id, font=('Helvetica', 12))
        self.search_button.grid(row=0, column=1, padx=10, pady=5, sticky='w')

        # Bind the click and focusout events
        self.search_entry.bind('<FocusIn>', self.on_entry_click)
        self.search_entry.bind('<FocusOut>', self.on_focusout)

        


        # Status dropdown
        self.status_var = StringVar(self.root)
        self.status_var.set("All Status")  # Default value

        self.status_menu = OptionMenu(search_frame, self.status_var, "All Status", "准确", "货物数量足够", "数量缺失/不匹配", "未获取")
        self.status_menu.config(font=('Helvetica', 13))
        self.status_menu.grid(row=1, column=0, padx=10, pady=5, sticky='w')

        self.filter_button = Button(search_frame, text="Filter", command=self.filter_by_status, font=('Helvetica', 13))
        self.filter_button.grid(row=1, column=1, padx=10, pady=5, sticky='w')






        # Action Buttons Frame
        action_frame = tk.Frame(main_frame)
        action_frame.grid(row=1, column=0, pady=10)

        self.load_button = Button(action_frame, text="Start Checking", command=self.start_checking, font=('Helvetica', 13))
        self.load_button.grid(row=0, column=0, padx=10)

        self.refresh_button = Button(action_frame, text="Refresh", command=self.refresh, font=('Helvetica', 13))
        self.refresh_button.grid(row=0, column=1, padx=10)

        self.export_button = Button(action_frame, text="Export Results", command=self.export_results, font=('Helvetica', 13))
        self.export_button.grid(row=0, column=2, padx=10)

        self.language_button = Button(action_frame, textvariable=self.language_var, command=self.change_language, font=('Helvetica', 13))
        self.language_button.grid(row=0, column=3, padx=10)




        # Status Text Area with horizontal and vertical scrollbars
        # Table display section
        table_frame = ttk.Frame(root)
        table_frame.grid(row=1, column=0, pady=0, padx=50, sticky='nsew')

        columns = ('产品代码', '产品名称', '预定数量', '包装数量', '最低单价', '最高单价', '价格差异', '状态')
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        # Define column headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        self.tree.grid(row=0, column=0, sticky='nsew')  # Ensure the treeview expands

        # Scrollbars for the table (Treeview widget)
        tree_scrollbar_y = tk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scrollbar_y.grid(row=0, column=1, sticky='ns')
        self.tree['yscrollcommand'] = tree_scrollbar_y.set

        



        # Configuring resizing for table and text areas
        root.grid_rowconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=3)
        root.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Add this within the __init__ method, after creating the table_frame
        summary_frame = tk.LabelFrame(root, text="Summary", padx=10, pady=10, font=('Helvetica', 18))
        summary_frame.grid(row=2, column=0, padx=50, pady=10, sticky='ew')

        # Add labels for displaying the summary
        self.matches_label = Label(summary_frame, text="Number of Matches 准确的值: 0", font=('Helvetica', 12))
        self.matches_label.grid(row=0, column=0, sticky='w', padx=10)

        self.sufficient_quantity_label = Label(summary_frame, text="Number of Sufficient Quantity 货物数量足够的值: 0", font=('Helvetica', 12))
        self.sufficient_quantity_label.grid(row=1, column=0, sticky='w', padx=10)

        self.mismatches_label = Label(summary_frame, text="Number of Mismatches 数量缺失/不匹配的值: 0", font=('Helvetica', 12))
        self.mismatches_label.grid(row=2, column=0, sticky='w', padx=10)

        self.not_available_label = Label(summary_frame, text="Number of Not Available 未获取的值: 0", font=('Helvetica', 12))
        self.not_available_label.grid(row=3, column=0, sticky='w', padx=10)

        self.loading_label = Label(root, text="", font=('Helvetica', 14), fg='red')  # Create an empty label
        self.loading_label.grid(row=1, column=0, pady=10)  # Adjust position as needed

    
    def change_language(self, *args):
        # Toggle between languages
        current_lang = self.language_var.get()
        if current_lang == "en":
            self.language_var.set("zh")  # Set to Mandarin
        else:
            self.language_var.set("en")  # Set to English

        print(f"Language changed to: {self.language_var.get()}")  # Debugging statement
        self.switch_language(self.language_var.get())



    #to change the language of the application 

    def switch_language(self, lang):
        print(f"Switching to language: {lang}")  # Debug statement
        if lang not in translations:
            print(f"Language {lang} not found in translations!")
            return
    
        # Update the UI texts based on the selected language
        self.root.title(translations[lang]["title"])
        self.select_order_files_button.config(text=translations[lang]["select_order_files"])
        self.select_packing_file_button.config(text=translations[lang]["select_packing_file"])
        self.search_button.config(text=translations[lang]["search"])
        self.status_var.set(translations[lang]["status_selection"])
        self.filter_button.config(text=translations[lang]["filter"])
        self.load_button.config(text=translations[lang]["start_checking"])
        self.refresh_button.config(text=translations[lang]["refresh"])
        self.export_button.config(text=translations[lang]["export_results"])
        self.back_button.config(text=translations[lang]["back to launcher"])

        
        # Update table column headers
        for idx, col in enumerate(translations[lang]["columns"]):
            self.tree.heading(f'#{idx+1}', text=col)


    def run_checking_process(self):
        self.show_loading()

        time.sleep(3)

        if not self.order_files or not self.packing_file:
                messagebox.showwarning("Warning", "Please upload both order and packing files.")
                self.hide_loading()
                return

        result_df = self.process_files()

        if isinstance(result_df, pd.DataFrame):
                # Display the results in the Treeview
                self.display_results(result_df)

                # Update the summary labels with actual values
                num_matches = len(result_df[result_df['状态'] == '准确'])
                num_sufficient = len(result_df[result_df['状态'] == '货物数量足够'])
                num_mismatches = len(result_df[result_df['状态'] == '数量缺失/不匹配'])
                num_not_available = len(result_df[result_df['状态'] == '未获取'])

                self.matches_label.config(text=f"Number of Matches: {num_matches}")
                self.sufficient_quantity_label.config(text=f"Number of Sufficient Quantity: {num_sufficient}")
                self.mismatches_label.config(text=f"Number of Mismatches: {num_mismatches}")
                self.not_available_label.config(text=f"Number of Not Available: {num_not_available}")
        else:
            messagebox.showerror("Error", "No valid data to display. Please check your files.")    

        self.hide_loading()  # Hide loading indicator after processing
        
        # Update these labels after processing the files in the start_checking method
    def start_checking(self):
        checking_thread = threading.Thread(target=self.run_checking_process)
        checking_thread.start()
            # if not self.order_files or not self.packing_file:
            #     messagebox.showwarning("Warning", "Please upload both order and packing files.")
            #     return

            # result_df = self.process_files()

            # if isinstance(result_df, pd.DataFrame):
            #     # Display the results in the Treeview
            #     self.display_results(result_df)

            #     # Update the summary labels with actual values
            #     num_matches = len(result_df[result_df['状态'] == '准确'])
            #     num_sufficient = len(result_df[result_df['状态'] == '货物数量足够'])
            #     num_mismatches = len(result_df[result_df['状态'] == '数量缺失/不匹配'])
            #     num_not_available = len(result_df[result_df['状态'] == '未获取'])

            #     self.matches_label.config(text=f"Number of Matches: {num_matches}")
            #     self.sufficient_quantity_label.config(text=f"Number of Sufficient Quantity: {num_sufficient}")
            #     self.mismatches_label.config(text=f"Number of Mismatches: {num_mismatches}")
            #     self.not_available_label.config(text=f"Number of Not Available: {num_not_available}")
            # else:
            #     messagebox.showerror("Error", "No valid data to display. Please check your files.")    

    def show_loading(self):
        self.loading_label.config(text="Processing... Please wait.")
        self.root.update_idletasks()  # Ensure the label updates immediately

    # Hide the loading indicator
    def hide_loading(self):
        self.loading_label.config(text="")
        self.root.update_idletasks()  # Ensure the label clears immediately

    def on_entry_click(self,event):
            if self.search_entry.get() == 'Type to search...':
                self.search_entry.delete(0, "end")  # Remove placeholder text
                self.search_entry.config(fg='black')  # Change text color to black

    def on_focusout(self, event):
            if self.search_entry.get() == '':
                self.search_entry.insert(0, 'Type to search...')  # Restore placeholder text
                self.search_entry.config(fg='grey')  # Change text color to grey
    
    def select_order_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if files:
            self.order_files = [file for file in files]
            self.update_file_labels()
            messagebox.showinfo("Success", "Order files successfully uploaded.")
    
    def select_packing_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file:
            self.packing_file = file
            self.update_file_labels()
            messagebox.showinfo("Success", "Packing file successfully uploaded.")
    
    def update_file_labels(self):
        if self.order_files:
            # Limit the number of displayed filenames to 3, followed by "and X more..." if necessary
            max_display_files = 3
            displayed_files = [file.split("/")[-1] for file in self.order_files[:max_display_files]]
            
            if len(self.order_files) > max_display_files:
                display_text = ", ".join(displayed_files) + f", and {len(self.order_files) - max_display_files} more..."
            else:
                display_text = ", ".join(displayed_files)
            
            self.order_files_label.config(text=f"Order Files: {display_text}")
        else:
            self.order_files_label.config(text="Order Files: None")

        if self.packing_file:
            # Display the name of the packing file
            packing_file_name = self.packing_file.split("/")[-1]
            self.packing_file_label.config(text=f"Packing File: {packing_file_name}")
        else:
            self.packing_file_label.config(text="Packing File: None")

    
    def process_files(self):
        try:
            # Load and process order files
            order_list = pd.concat([pd.read_excel(file, skiprows=12) for file in self.order_files], ignore_index=True)
            order_list.rename(columns={'品号': '产品代码', '数量': '预定数量', '品名': '产品名称', '单价': '单价'}, inplace=True)
            order_list = self.clean_and_convert_to_float(order_list, '预定数量')
            order_list = self.clean_and_convert_to_float(order_list, '单价')

            # Calculate the highest and lowest unit price for each product
            price_summary = order_list.groupby('产品代码')['单价'].agg(['min', 'max']).reset_index()
            price_summary.rename(columns={'min': '最低单价', 'max': '最高单价'}, inplace=True)

            # Calculate the price difference and add it to the price_summary DataFrame
            price_summary['价格差异'] = price_summary['最高单价'] - price_summary['最低单价']

            # Sum the quantities for each product
            order_list_sum = order_list.groupby('产品代码')['预定数量'].sum().reset_index()
            order_list_sum = pd.merge(order_list_sum, order_list[['产品代码', '产品名称']].drop_duplicates(), on='产品代码', how='left')
            order_list_sum = pd.merge(order_list_sum, price_summary, on='产品代码', how='left')

            # Load the packing list
            packing_list = pd.read_excel(self.packing_file, skiprows=6)
            packing_list.rename(columns={'Unnamed: 6': '包装数量', '物料号': '物料号', 'Materials No.': '物料号'},inplace=True)
            packing_list = self.clean_and_convert_to_float(packing_list, '包装数量')
            packing_list = self.split_and_expand_packing_list(packing_list, '物料号', '包装数量')
            packing_list_sum = packing_list.groupby('物料号')['包装数量'].sum().reset_index()

            # Merge the order list and packing list
            merged_df = pd.merge(order_list_sum, packing_list_sum, left_on='产品代码', right_on='物料号', how='left')

            # Define the status column based on comparison of quantities
            merged_df['状态'] = '准确'
            merged_df.loc[merged_df['包装数量'].isna() | (merged_df['包装数量'] == 0), '状态'] = '未获取'
            merged_df.loc[merged_df['包装数量'] > merged_df['预定数量'], '状态'] = '货物数量足够'
            merged_df.loc[merged_df['包装数量'] < merged_df['预定数量'], '状态'] = '数量缺失/不匹配'

            # Select the relevant columns for output
            self.result_df = merged_df[['产品代码', '产品名称', '预定数量', '包装数量', '最低单价', '最高单价','价格差异', '状态']]

            price_diff_df = price_summary[price_summary['最低单价'] != price_summary['最高单价']]


            # Prepare results for display
            accurate_and_sufficient_df = self.result_df[self.result_df['状态'].isin(['准确', '货物数量足够'])]
            mismatch_df = self.result_df[self.result_df['状态'] == '数量缺失/不匹配']
            unavailable_df = self.result_df[self.result_df['状态'] == '未获取']

            summary = []
            summary.append(f"\nTable 1: Accurate and Sufficient Quantity 准确而充足的数量:\n{accurate_and_sufficient_df.to_string(index=False)}")
            summary.append(f"\nTable 2: Quantity Mismatch/Shortage 数量不匹配/短缺:\n{mismatch_df.to_string(index=False)}")
            summary.append(f"\nTable 3: Not Available 未获取:\n{unavailable_df.to_string(index=False)}")
            summary.append(f"\nTable 4: Products with Different Prices 价格不同的产品:\n{price_diff_df.to_string(index=False)}")
            summary.append(f"\nNumber of Matches 准确的值: {len(accurate_and_sufficient_df[accurate_and_sufficient_df['状态'] == '准确'])}")
            summary.append(f"Number of Sufficient Quantity 货物数量足够的值: {len(accurate_and_sufficient_df[accurate_and_sufficient_df['状态'] == '货物数量足够'])}")
            summary.append(f"Number of Mismatches 数量缺失/不匹配的值: {len(mismatch_df)}")
            summary.append(f"Number of Not Available 未获取的值: {len(unavailable_df)}")

            return self.result_df

        except Exception as e:
            # In case of an error, return None
            messagebox.showerror("Error", f"Error processing files: {str(e)}")
            return None

    
    def clean_and_convert_to_float(self, df, column_name):
        # Remove non-numeric characters (like commas, spaces)
        df[column_name] = df[column_name].astype(str).replace(r'[^\d.]', '', regex=True)
        # Convert to numeric, forcing errors to NaN
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce')
        return df
    
    def split_and_expand_packing_list(self, df, id_column, quantity_column):
        # Split the '物料号' by '/' if combined product IDs are present
        expanded_df = df[id_column].str.split('/', expand=True).stack().reset_index(level=1, drop=True)
        expanded_df.name = '物料号'
        
        # Join with the original DataFrame, duplicating quantities for each split ID
        df = df.drop(columns=[id_column]).join(expanded_df)
        
        return df
    
    def display_results(self, result_df):
    # Clear previous data from the Treeview
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Insert new data into the Treeview
        for _, row in result_df.iterrows():
            self.tree.insert("", "end", values=list(row))

    
    def search_product_id(self):
        search_id = self.search_entry.get().strip()
        if self.result_df is not None and search_id:
            search_df = self.result_df[self.result_df['产品代码'].astype(str).str.contains(search_id)]
            if not search_df.empty:
                self.display_results(search_df)  # Update Treeview with search results
            else:
                messagebox.showinfo("Search Result", f"No results found for '{search_id}'")
        else:
            messagebox.showwarning("Search Error", "Please enter a product ID to search or ensure results are available.")

    
    def filter_by_status(self):
        status_filter = self.status_var.get()
        if self.result_df is not None:
            if status_filter == "All Status":
                filtered_df = self.result_df
            else:
                filtered_df = self.result_df[self.result_df['状态'] == status_filter]
            
            # Display the filtered results (pass the DataFrame directly, not its string representation)
            self.display_results(filtered_df)
        else:
            messagebox.showwarning("Filter Error", "No results available to filter.")

    
    def refresh(self):
        confirm = messagebox.askyesno("Confirm Refresh", "Are you sure you want to refresh and clear all data?")
        if confirm:
            self.order_files = []
            self.packing_file = None
            self.result_df = None
            self.update_file_labels()
            messagebox.showinfo("Refreshed", "All data has been cleared.")
            self.status_text.delete(1.0, tk.END)
            

    def export_results(self):
        if self.result_df is None:
            messagebox.showwarning("Warning", "No results available to export.")
            return
    
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # Prepare results
                accurate_and_sufficient_df = self.result_df[self.result_df['状态'].isin(['准确', '货物数量足够'])]
                mismatch_df = self.result_df[self.result_df['状态'] == '数量缺失/不匹配']
                unavailable_df = self.result_df[self.result_df['状态'] == '未获取']

                # Create a new Excel writer
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # Write the three tables to separate sheets
                    accurate_and_sufficient_df.to_excel(writer, sheet_name='Accurate and Sufficient', index=False)
                    mismatch_df.to_excel(writer, sheet_name='Mismatch', index=False)
                    unavailable_df.to_excel(writer, sheet_name='Not Available', index=False)

                    # Prepare the summary data
                    summary_data = {
                        'Status': ['准确', '货物数量足够', '数量缺失/不匹配', '未获取'],
                        'Count': [
                            len(accurate_and_sufficient_df[accurate_and_sufficient_df['状态'] == '准确']),
                            len(accurate_and_sufficient_df[accurate_and_sufficient_df['状态'] == '货物数量足够']),
                            len(mismatch_df),
                            len(unavailable_df)
                        ]
                    }
                    summary_df = pd.DataFrame(summary_data)

                    # Write the summary to a separate sheet
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
                messagebox.showinfo("Success", f"Results successfully exported to {file_path}.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export results: {str(e)}")

    
    def back_to_launcher(self):
        """Closes the current window and reopens the startup window."""
        self.root.destroy()  # Close the current window
        Esmart_ACS()  # Call the startup function to relaunch the startup window


    # def return_to_startup(self):
    #     self.root.destroy()  # Close the current window
    #     self.StartupApp()  # Show the startup window again

# Language dictionary
translations = {
    "en": {
        "files_frame" : "File Selection",
        "title": "Order and Packing List Checker Newmen Group Internal Order",
        "select_order_files": "Select Order Files",
        "select_packing_file": "Select Packing File",
        "order_files": "Order Files: ",
        "packing_file": "Packing File: ",
        "search": "Search",
        "status_selection": "Status Selection",
        "filter": "Filter",
        "start_checking": "Start Checking",
        "refresh": "Refresh",
        "export_results": "Export Results",
        "status": "Status",
        "columns": ['Product Code', 'Product Name', 'Order Quantity', 'Packed Quantity', 'Min Price', 'Max Price', 'Price Difference', 'Status'],
        "back to launcher":"Back to launcher",


    },
    "zh": {
        "files_frame" : "文件选择",
        "title": "订单与包装清单检查 新贵集团Newmen 内部订单",
        "select_order_files": "选择订单文件",
        "select_packing_file": "选择包装文件",
        "order_files": "订单文件: ",
        "packing_file": "包装文件: ",
        "search": "搜索",
        "status_selection": "状态选择",
        "filter": "筛选",
        "start_checking": "开始检查",
        "refresh": "刷新",
        "export_results": "导出结果",
        "status": "状态",
        "columns": ['产品代码', '产品名称', '预定数量', '包装数量', '最低单价', '最高单价', '价格差异', '状态'],
        "back to launcher":"回到启动器",

    }
}

        

if __name__ == "__main__":
    root = tk.Tk()
    app = CheckingApp(root)
    root.mainloop()
