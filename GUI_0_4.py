import pandas as pd
import re
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from screeninfo import get_monitors
import logging
from openpyxl.utils import get_column_letter

#testchange

#testchange2

# Setup logging
logging.basicConfig(filename='app.log', level=logging.INFO)

class ExcelProcessor:
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.output_path = output_path

    @staticmethod
    def extract_year(row):
        if pd.isna(row) or not isinstance(row, str):
            return ''
        matches_years = re.findall(r'\b\d{4}\b', row)
        return matches_years[0] if matches_years else ''

    @staticmethod
    def extract_numbers(row):
        if pd.isna(row) or not isinstance(row, str):
            return ''
        matches_numbers = re.findall(r'(?<!\()\b\d{2}\b(?!\.)', row)
        return matches_numbers[0] if matches_numbers else ''
    
    def process_file(self):
        try:
            wb = pd.read_excel(self.input_path)
            wb['Softwarebezeichnung'] = wb['Softwarebezeichnung'].apply(self.shorten_title)

            #filter out empty entries in "Softwarebezeichnung"/Erasing bad data generated
            wb = wb[wb['Softwarebezeichnung'].str.strip().astype(bool)]

            wb["TransformedInstallationsanzahl"] = wb.apply(
                lambda row: 
                self.extract_year(row['Softwarebezeichnung']) + 
                (": " if self.extract_year(row['Softwarebezeichnung']) else "") + 
                self.extract_numbers(row['Softwarebezeichnung']) +
                (": " if self.extract_numbers(row['Softwarebezeichnung']) else "") +  
                str(row['Installationsanzahl']) + "x" + " (" + str(row["Version"]) + ")" if not pd.isna(row["Version"]) else "",
                axis=1
            )
            
            aggregated_data = wb.groupby('Softwarebezeichnung')['TransformedInstallationsanzahl'].agg(', '.join).reset_index()
            total_installations = wb.groupby('Softwarebezeichnung')['Installationsanzahl'].sum()
            aggregated_data['Gesamtanzahl'] = total_installations.loc[aggregated_data['Softwarebezeichnung']].astype(str).values
            aggregated_data['Version Details'] = aggregated_data["TransformedInstallationsanzahl"].astype(str)

            logging.info('File processed successfully')
            return aggregated_data
        
        except Exception as e:
            logging.error(f'Error processing file: {e}')
            raise

    @staticmethod
    def shorten_title(row):
        if pd.isna(row) or not isinstance(row, str):
            return ''
        # Remove Version numbers
        title_without_versions = re.sub(r'\b\d+(\.\d+){1,}\b', '', row)
        # Remove exactly 2 digits (e.g. Flash Player 30)
        title_without_two_digits = re.sub(r'(?<!\()\b\d{2}\b', '', title_without_versions)
        # Remove exactly 4 digits in a typical year format
        title_without_year = re.sub(r'\b(19|20)\d{2}\b', '', title_without_two_digits)
        # Remove everything in brackets (e.g. (64-Bit)
        title_without_bracket = re.sub(r'\(.*', '', title_without_year)
        # Remove everything after a hyphen for clearing lists
        title_without_hyphen = re.sub(r'\s\-(.*)', '', title_without_bracket)
        # Remove exactly 4 digits
        title_without_4_digits = re.sub(r'\b\d{4}\b', '', title_without_hyphen)
        # Remove Capital V if followed by a digit without a space and everything after it
        title_without_v = re.sub(r'V\d{1}.*', '', title_without_4_digits) 
        # Remove "B" followed by 3 digits
        title_without_B = re.sub(r'B\d{3}', '', title_without_v)
        # Remove everything after "Kit"
        title_without_Kit = re.sub(r'Kit.*', 'Kit', title_without_B)
        # Remove exactly 3 digits
        title_without_3_digits = re.sub(r'\b\d{3}\b', '', title_without_Kit)
        # Replace IV followed by anything with IV
        title_without_IV = re.sub(r'IV.*', 'IV', title_without_3_digits)
        # ELAN Touchpad 15.12.1.3 X64 WHOL becomes ELAN Touchpad
        title_without_version = re.sub(r'15.*', '', title_without_IV)
        # Remove titles consisting of one or more dots
        title_shortened=re.sub(r'^[.\s]+$', '', title_without_version)
        # Return shortened title
        return title_shortened.strip()
    
class ApplicationGUI:
    def __init__(self, root):
        self.root = root
        self.sorting_order = {}
        self.setup_gui()

    def setup_gui(self):
        self.root.title("Software-Inventarisierung BETA")
        monitors = get_monitors()
        main_monitor = min(monitors, key=lambda monitor: monitor.x)
        screen_width = main_monitor.width // 2
        screen_height = main_monitor.height // 2
        self.root.geometry(f'{screen_width}x{screen_height}+0+0')
        label = tk.Label(self.root, text='Drag and Drop your file here', relief='raised')
        label.pack(fill=tk.BOTH, expand=1)
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop)

    def drop(self, event):
        file_input = event.data.strip("{}")
        logging.info(f'File dropped: {file_input}')
        processor = ExcelProcessor(file_input, '')
        try:
            processed_data = processor.process_file()
            # Select only the desired columns
            final_data = processed_data[['Softwarebezeichnung', 'Gesamtanzahl', 'Version Details']]
            self.display_data_in_table(final_data)
            self.close_dnd_box()
        except Exception as e:
            logging.error(f'Error during file processing: {e}')
            self.show_error_message(f'Error: {e}')

    def close_dnd_box(self):
        self.root.withdraw()

    def show_error_message(self, message):
        error_window = tk.Toplevel(self.root)
        error_window.title("Error")
        tk.Label(error_window, text=message).pack()
        tk.Button(error_window, text="OK", command=error_window.destroy).pack()

    def on_results_window_close(self):
        self.root.destroy()
        
    def treeview_sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        try:
            l.sort(key=lambda t: int(t[0]), reverse=reverse)  # Assume column values are integers
        except ValueError:
            l.sort(reverse=reverse)  # Fallback to string sorting

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        # Reverse sort next time
        self.sorting_order[col] = not reverse
        self.tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))

    def apply_filter(self):
        filter_text = self.filter_var.get().lower()
        self.tree.delete(*self.tree.get_children())  # Clear the current items in the treeview

        # Filter the data at the DataFrame level
        filtered_data = self.original_data[self.original_data['Softwarebezeichnung'].str.lower().str.contains(filter_text)]

       # Insert the filtered data into the treeview
        for _, row in filtered_data.iterrows():
            self.tree.insert('', 'end', values=row.tolist())

    def display_data_in_table(self, data):
        table_window = tk.Toplevel(self.root)
        table_window.title("Processed Data")

        # Frame for filter widgets
        filter_frame = ttk.Frame(table_window)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        # Filter Entry
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda name, index, mode: self.apply_filter())
        filter_entry = ttk.Entry(filter_frame, textvariable=self.filter_var)
        filter_entry.pack(side=tk.LEFT, padx=(0, 10))

        # Filter Button
        filter_button = ttk.Button(filter_frame, text="Filter", command=self.apply_filter)
        filter_button.pack(side=tk.LEFT)

        # Treeview Frame
        tree_frame = ttk.Frame(table_window)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Setup the Treeview widget with scrollbars
        self.tree = ttk.Treeview(tree_frame, columns=data.columns.tolist(), show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Pack the Treeview and Scrollbars
        #self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        #vsb.pack(side=tk.RIGHT, fill=tk.Y)
        #hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Arrange the Treeview and Scrollbars using grid layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # Configure the grid to expand properly when the window is resized
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Setup columns, heading and sorting
        for col in data.columns:
            self.sorting_order[col] = True
            self.tree.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(_col, self.sorting_order[_col]))

        # Get the width of the window
        window_width = table_window.winfo_screenwidth()

        # Set the column width as a fraction of the window width
        if col == 'Softwarebezeichnung':
            self.tree.column(col, width=int(window_width * 0.2), minwidth=int(window_width * 0.2))  # 20% of the window width
        elif col == 'Gesamtanzahl':
            self.tree.column(col, width=int(window_width * 0.1), minwidth=int(window_width * 0.1))  # 10% of the window width
        elif col == 'Version Details':
            self.tree.column(col, width=int(window_width * 0.7), minwidth=int(window_width * 0.7))  # 70% of the window width
            
        # Populate the Treeview with data
        self.original_data = data  # Store the original data
        for _, row in data.iterrows():
            self.tree.insert('', 'end', values=row.tolist())

        self.tree.tag_configure('filtered', background='lightyellow')  # Optional: Highlight filtered rows

        # Store original data for filtering
        self.original_data = data

        # Display the window in full screen
        table_window.attributes('-zoomed', True)
        
        table_window.protocol("WM_DELETE_WINDOW", self.on_results_window_close)

if __name__ == '__main__':
    try:
        root = TkinterDnD.Tk()
        app = ApplicationGUI(root)
        app.root.mainloop()
    except Exception as e:
        logging.error(f'Application failed to start: {e}')
