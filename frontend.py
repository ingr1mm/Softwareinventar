import tkinter as tk
from tkinter import ttk
import windnd
import logging
from screeninfo  import get_monitors
from backend import ExcelProcessor

class ApplicationGUI:
    def __init__(self, root, db_config):
        self.root = root
        self.db_config = db_config        
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

        # Windows-DND
        windnd.hook_dropfiles(label, func=self.drop)

    def drop(self, files):
        file_path = files[0].decode('utf-8')  # Decode the first file from bytes to string
        logging.info(f'File dropped: {file_path}')
        processor = ExcelProcessor(file_path, self.db_config)
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

        # Horizontal scrollbar
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x')

        # Populate the Treeview with data
        for _, row in data.iterrows():
            self.tree.insert('', 'end', values=(row['Softwarebezeichnung'], row['Softwarekategorie'], row['Fachbereich'], row['Softwarebeschreibung'],
                                                row['Gesamtanzahl'], row['Version Details']))

        # Store original data for filtering
        self.original_data = data

        # Highlight filtered rows 
        self.tree.tag_configure('filtered', background='lightyellow')
        
        # Closing all windows
        table_window.protocol("WM_DELETE_WINDOW", self.on_results_window_close) 