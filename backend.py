import pandas as pd
import re
import logging
import mysql.connector
from config import db_config


class ExcelProcessor:
    def __init__(self, input_path, db_config):
        self.input_path = input_path
        self.db_config = db_config
        
    def fetch_software_info(self, software_name):
        if pd.isna(software_name) or software_name is None:
                return ("", "", "")  

        try:
                connection = mysql.connector.connect(**self.db_config)
                cursor = connection.cursor()
                query = "SELECT Softwarekategorie, Fachbereich, Softwarebeschreibung FROM Softwareinformationen WHERE Softwarebezeichnung = %s"
                cursor.execute(query, (software_name,))
                result = cursor.fetchone()
                connection.close()
                return result if result else ("", "", "")
        except mysql.connector.Error as err:
                logging.error(f"Error: {err}")
                return ("", "", "")
              

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