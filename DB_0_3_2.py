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
        title_shortened_specials=re.sub(r'^[.\s]+$', '', title_without_version)
        # Remove titles consisting of 2 or less characters
        title_shortened=re.sub(r'^.{0,2}$', '', title_shortened_specials)
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

    def on_treeview_hover(self, event):
        # Identify the row and column hovered over
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)
            # Get the value of the cell and show it as a tooltip
            # Adjusting column index to accommodate treeview column indexing starting at #1
            value = self.tree.item(row_id, 'values')[int(column.replace('#', '')) - 1]
            self.tooltip.show_tip(value)
        else:
            self.tooltip.hide_tip()

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

    def apply_filterBezeichnung(self):
        filter_text = self.filter_varBezeichnung.get().lower()
        self.tree.delete(*self.tree.get_children())  # Clear the current items in the treeview

        # Filter the data at the DataFrame level
        filtered_data = self.original_data[self.original_data['Softwarebezeichnung'].str.lower().str.contains(filter_text)]

       # Insert the filtered data into the treeview
        for _, row in filtered_data.iterrows():
            self.tree.insert('', 'end', values=row.tolist())

    def apply_filterKategorie(self):
        filter_text = self.filter_varKategorie.get().lower()
        self.tree.delete(*self.tree.get_children())  # Clear the current items in the treeview

        # Filter the data at the DataFrame level
        filtered_data = self.original_data[self.original_data['Softwarekategorie'].str.lower().str.contains(filter_text)]

       # Insert the filtered data into the treeview
        for _, row in filtered_data.iterrows():
            self.tree.insert('', 'end', values=row.tolist())

    def apply_filterFachbereich(self):
        filter_text = self.filter_varFachbereich.get().lower()
        self.tree.delete(*self.tree.get_children())  # Clear the current items in the treeview

        # Filter the data at the DataFrame level
        filtered_data = self.original_data[self.original_data['Fachbereich'].str.lower().str.contains(filter_text)]

       # Insert the filtered data into the treeview
        for _, row in filtered_data.iterrows():
            self.tree.insert('', 'end', values=row.tolist())

    def display_data_in_table(self, data):
        table_window = tk.Toplevel(self.root)
        table_window.title("Processed Data")
        
        # Erstellen eines Containers für Filter
        filter_container = ttk.Frame(table_window)
        filter_container.pack(fill=tk.X, padx=10, pady=5)

        # Filter für Softwarebezeichnung
        label_bezeichnung = ttk.Label(filter_container, text="Bezeichnung:")
        label_bezeichnung.grid(row=0, column=0, padx=(0, 10), sticky='w')

        self.filter_varBezeichnung = tk.StringVar()
        filter_entryBezeichnung = ttk.Entry(filter_container, textvariable=self.filter_varBezeichnung)
        filter_entryBezeichnung.grid(row=1, column=0, padx=(0, 10), sticky='we')
        self.filter_varBezeichnung.trace_add("write", lambda name, index, mode: self.apply_filterBezeichnung())

        # Filter für Softwarekategorie
        label_kategorie = ttk.Label(filter_container, text="Kategorie:")
        label_kategorie.grid(row=0, column=1, padx=(0, 10), sticky='w')

        self.filter_varKategorie = tk.StringVar()
        filter_entryKategorie = ttk.Entry(filter_container, textvariable=self.filter_varKategorie)
        filter_entryKategorie.grid(row=1, column=1, padx=(0, 10), sticky='we')
        self.filter_varKategorie.trace_add("write", lambda name, index, mode: self.apply_filterKategorie())

        # Filter für Fachbereich
        label_fachbereich = ttk.Label(filter_container, text="Fachbereich:")
        label_fachbereich.grid(row=0, column=2, padx=(0, 10), sticky='w')

        self.filter_varFachbereich = tk.StringVar()
        filter_entryFachbereich = ttk.Entry(filter_container, textvariable=self.filter_varFachbereich)
        filter_entryFachbereich.grid(row=1, column=2, padx=(0, 10), sticky='we')
        self.filter_varFachbereich.trace_add("write", lambda name, index, mode: self.apply_filterFachbereich())

        # Einstellung, damit sich die Entry-Widgets horizontal ausdehnen
        filter_container.grid_columnconfigure(0, weight=1)
        filter_container.grid_columnconfigure(1, weight=1)
        filter_container.grid_columnconfigure(2, weight=1)

        # Setup the Treeview widget within a frame for better layout control
        tree_frame = ttk.Frame(table_window)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Define columns for the Treeview
        columns = ['Softwarebezeichnung', 'Softwarekategorie', 'Fachbereich', 'Softwarebeschreibung', 'Gesamtanzahl', 'Version Details']
        
        # Create the Treeview widget
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='w', minwidth=100)
            if col == 'Gesamtanzahl':
                self.tree.column('Gesamtanzahl', anchor='center', width=50)
                pass          
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Vertical scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')

        # Tooltip for the Treeview
        self.tooltip = ToolTip(self.tree)

        # Bind the hover event to the treeview
        self.tree.bind("<Motion>", self.on_treeview_hover)
        self.tree.bind("<Leave>", lambda e: self.tooltip.hide_tip())

        # Populate the Treeview with data
        for _, row in data.iterrows():
            self.tree.insert('', 'end', values=(row['Softwarebezeichnung'], row['Softwarekategorie'], row['Fachbereich'], row['Softwarebeschreibung'],
                                                row['Gesamtanzahl'], row['Version Details']))

        # Store original data for filtering
        self.original_data = data

        # Highlight filtered rows 
        self.tree.tag_configure('filtered', background='lightyellow')

        # Display option
        table_window.attributes('-zoomed', True)
        
        # Closing all windows
        table_window.protocol("WM_DELETE_WINDOW", self.on_results_window_close)  
 
class ToolTip(object):
    def __init__(self, tree):
        self.widget = tree
        self.tip_window = None

    def show_tip(self, text):
        if self.tip_window or not text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


if __name__ == '__main__':
    try:
        root = TkinterDnD.Tk()
        app = ApplicationGUI(root)
        app.root.mainloop()
    except Exception as e:
        logging.error(f'Application failed to start: {e}')


