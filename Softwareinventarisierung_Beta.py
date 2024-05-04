import pandas as pd
import re
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from screeninfo import get_monitors
import logging
import mysql.connector

# Setup logging
logging.basicConfig(filename='app.log', level=logging.INFO)

#db-connection data
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',  
    'database': 'Softwarebestand',
    'port': 3306
}

class ExcelProcessor:
    def __init__(self, input_path, db_config):
        self.input_path = input_path
        self.db_config = db_config
        
    def fetch_software_info(self, software_name):
        if pd.isna(software_name) or software_name is None:
                return ("", "")  # Return default values or handle as needed

        try:
                connection = mysql.connector.connect(**db_config)
                cursor = connection.cursor()
                query = "SELECT Softwarekategorie, Fachbereich, Softwarebeschreibung FROM Softwareinformationen WHERE Softwarebezeichnung = %s"
                cursor.execute(query, (software_name,))
                result = cursor.fetchone()
                connection.close()
                return result if result else ("", "")
        except mysql.connector.Error as err:
                logging.error(f"Error: {err}")
                return ("", "")

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

            # Fetch software information from the database
            for index, row in wb.iterrows():                
                software_name = row['Softwarebezeichnung']
                kategorie, fachbereich, beschreibung = self.fetch_software_info(software_name)
                #if kategorie and fachbereich and beschreibung:
                wb.at[index, 'Softwarekategorie'] = kategorie if kategorie else ""
                wb.at[index, 'Fachbereich'] = fachbereich if fachbereich else ""  
                wb.at[index, 'Softwarebeschreibung'] = beschreibung if beschreibung else "" 

            # Adjusted aggregation logic
            aggregated_data = wb.groupby('Softwarebezeichnung', as_index=False).agg({
                'TransformedInstallationsanzahl': ', '.join, 
                'Installationsanzahl': 'sum',
                'Softwarekategorie': 'first',  
                'Fachbereich': 'first',
                'Softwarebeschreibung': 'first'

            }).rename(columns={'Installationsanzahl': 'Gesamtanzahl',
                               'TransformedInstallationsanzahl': 'Version Details'})
                    
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
        processor = ExcelProcessor(file_input, db_config)
        try:
            processed_data = processor.process_file()
            # Select only the desired columns
            final_data = processed_data[['Softwarebezeichnung', 'Softwarekategorie', 'Fachbereich', 'Softwarebeschreibung', 'Gesamtanzahl', 'Version Details']]
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

        # Setup the Treeview widget within a frame for better layout control
        tree_frame = ttk.Frame(table_window)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Define columns for the Treeview
        columns = ['Softwarebezeichnung', 'Softwarekategorie', 'Fachbereich', 'Softwarebeschreibung', 'Gesamtanzahl', 'Version Details']
        
        # Create the Treeview widget
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center', width=100)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Vertical scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')

        # Horizontal scrollbar
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x', after=self.tree)

        # Populate the Treeview with data
        for _, row in data.iterrows():
            self.tree.insert('', 'end', values=(row['Softwarebezeichnung'], row['Softwarekategorie'], row['Fachbereich'], row['Softwarebeschreibung'],
                                                row['Gesamtanzahl'], row['Version Details']))

        # Store original data for filtering
        self.original_data = data

        # Optional: Highlight filtered rows (if applicable)
        self.tree.tag_configure('filtered', background='lightyellow')

        # Fullscreen option
        table_window.attributes('-zoomed', True)

        # table_window.attributes('-fullscreen', True)  # Uncomment if you want to display the window in fullscreen
        
        table_window.protocol("WM_DELETE_WINDOW", self.on_results_window_close)   

if __name__ == '__main__':
    try:
        root = TkinterDnD.Tk()
        app = ApplicationGUI(root)
        app.root.mainloop()
    except Exception as e:
        logging.error(f'Application failed to start: {e}')