import os
import re
import pandas as pd
import pdfplumber
import docx
from odf import text, teletype
from odf.opendocument import load
from pathlib import Path
import logging
import subprocess
import tempfile
import shutil
import time

# --- Configuration ---
# Configure logging to see what happens in the console
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

class ContactExtractor:
    def __init__(self, root_dir):
        self.root_dir = root_dir
        self.data_list = []
        self.failed_count = 0
        
        # Regex for French phone numbers (flexible: 06, +33, spaces, dots)
        self.phone_pattern = re.compile(r'(?:(?:\+|00)33|0)\s*[1-9](?:[\s.-]*\d{2}){4}')
        # Regex for Emails
        self.email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
        
        self.word_app = None
        if WIN32_AVAILABLE:
            try:
                self.word_app = win32.Dispatch("Word.Application")
                self.word_app.Visible = False # Word is running in background
                self.word_app.DisplayAlerts = False 
            except Exception as e:
                logger.warning(f"Impossible to launch Word : {e}")

    def run(self):
        """Main execution method."""
        logger.info(f"Starting scan in: {self.root_dir}")
        self.start_time = time.time()
        
        # Walk through directory recursively
        for root, dirs, files in os.walk(self.root_dir):
            for file in files:
                if not file.startswith('~$'): 
                    file_path = os.path.join(root, file)
                    self.process_file(file_path)
                    
        self.close_word()

        if not self.data_list:
            logger.warning("No data extracted. Check your files or table structures.")
            return

        self.save_data()

    def process_file(self, file_path):
        """Dispatcher that calls the right extractor based on extension."""
        ext = os.path.splitext(file_path)[1].lower()
        raw_texts_list = [] # On s'attend maintenant à une liste

        try:
            if ext == '.pdf':
                raw_texts_list = self.extract_from_pdf(file_path)
            elif ext == '.docx':
                raw_texts_list = self.extract_from_docx(file_path)
            elif ext == '.odt':
                raw_texts_list = self.extract_from_odt(file_path)
            elif ext == '.doc':
                raw_texts_list = self.extract_from_doc(file_path)
            else:
                return # Ignore other files
            
            # Si on a trouvé des données (liste non vide)
            if raw_texts_list:
                logger.info(f"Data found in: {os.path.basename(file_path)} ({len(raw_texts_list)} contacts potential)")
                
                # On boucle sur CHAQUE texte trouvé (chaque cellule est un contact potentiel)
                for text_blob in raw_texts_list:
                    # On ignore les cellules trop vides ou parasites (moins de 5 chars par ex)
                    if len(text_blob) < 5: 
                        continue
                        
                    parsed_info = self.parse_contact_info(text_blob)
                    parsed_info['source_file'] = file_path 
                    self.data_list.append(parsed_info)
            else:
                logger.debug(f"No valid table data found in {os.path.basename(file_path)}")
                self.failed_count += 1

        except Exception as e:
            logger.error(f"Error processing {file_path}: {e}")
            self.failed_count += 1
            
    def extract_from_doc(self, file_path):
        """
        Alternative: Converts .doc to .docx using LibreOffice, 
        then reads it with python-docx.
        """
        # Check if LibreOffice is available in PATH (usually 'soffice')
        # On Windows, you might need to add the full path, e.g.:
        # soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        soffice_path = shutil.which("soffice") 
        
        # Fallback for common Windows install location if not in PATH
        if not soffice_path:
            potential_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
            if os.path.exists(potential_path):
                soffice_path = potential_path

        if not soffice_path:
            logger.warning("LibreOffice (soffice) not found. Cannot process .doc file.")
            return None

        try:
            # Create a temporary directory to store the converted .docx
            with tempfile.TemporaryDirectory() as temp_dir:
                # Command to convert .doc to .docx
                cmd = [
                    soffice_path,
                    '--headless',
                    '--convert-to', 'docx',
                    '--outdir', temp_dir,
                    file_path
                ]
                
                # Run conversion (suppress output)
                subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
                
                # The file will have the same name but .docx extension
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                converted_file = os.path.join(temp_dir, base_name + ".docx")
                
                if os.path.exists(converted_file):
                    # Reuse your existing .docx logic!
                    return self.extract_from_docx(converted_file)
                else:
                    logger.warning(f"Conversion failed for {file_path}")
                    return None
                    
        except Exception as e:
            logger.error(f"Error converting .doc file: {e}")
            return None

    def extract_from_pdf(self, file_path):
        """Extracts text from all cells in the first 6 tables of a PDF."""
        extracted_texts = []
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return []
            first_page = pdf.pages[0]
            # On prend les 6 premiers tableaux
            tables = first_page.extract_tables()[:6] 
            
            for table in tables:
                # On aplatit le tableau (liste de listes -> liste simple)
                # et on ne garde que les cellules qui ont du texte
                for row in table:
                    for cell in row:
                        if cell and cell.strip():
                            extracted_texts.append(cell.strip())
                            
        return extracted_texts

    def extract_from_docx(self, file_path):
        """Extracts text from all cells in the first 6 tables of a DOCX."""
        extracted_texts = []
        doc = docx.Document(file_path)
        
        # On prend jusqu'à 6 tableaux
        target_tables = doc.tables[:6]
        
        for table in target_tables:
            for row in table.rows:
                for cell in row.cells:
                    text_content = cell.text.strip()
                    if text_content:
                        extracted_texts.append(text_content)
                        
        return extracted_texts

    def extract_from_odt(self, file_path):
        """Extracts text from all cells in the first 6 tables of an ODT."""
        extracted_texts = []
        doc = load(file_path)
        all_tables = doc.getElementsByType(text.Table)
        
        # On prend les 6 premiers
        target_tables = all_tables[:6]
        
        for table in target_tables:
            rows = table.getElementsByType(text.TableRow)
            for row in rows:
                cells = row.getElementsByType(text.TableCell)
                for cell in cells:
                    cell_text = teletype.extractText(cell).strip()
                    if cell_text:
                        extracted_texts.append(cell_text)
                        
        return extracted_texts

    def parse_contact_info(self, raw_text):
        """
        Parses raw text to identify names, phones, emails and addresses.
        Returns a flat dictionary with dynamic keys (phone_1, phone_2, etc.)
        """
        info = {}
        
        # 1. Clean up text lines
        lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
        
        # 2. Extract Emails
        emails = self.email_pattern.findall(raw_text)
        for i, email in enumerate(emails, 1):
            info[f'email_{i}'] = email
            
        # 3. Extract Phones
        phones = self.phone_pattern.findall(raw_text)
        for i, phone in enumerate(phones, 1):
            info[f'phone_{i}'] = phone

        # 4. Heuristic for Name and Address (CORRIGÉ)
        remaining_lines = []
        
        for line in lines:
            clean_line = line
            
            # Au lieu de vérifier si la ligne EST un téléphone, 
            # on SUPPRIME le téléphone de la ligne
            for p in phones:
                clean_line = clean_line.replace(p, "")
            
            # Idem pour les emails
            for e in emails:
                clean_line = clean_line.replace(e, "")
            
            # On nettoie les espaces multiples qui pourraient rester (ex: "Paris  ")
            clean_line = clean_line.strip()
            
            # S'il reste du texte après avoir enlevé emails et téléphones, c'est une partie de l'adresse/nom
            # On ignore les lignes qui deviennent vides ou qui ne contiennent que des caractères parasites
            if len(clean_line) > 1: 
                remaining_lines.append(clean_line)

        # Assume First remaining line is the Name/Company
        if remaining_lines:
            info['name'] = remaining_lines[0]
            # Rest is address
            if len(remaining_lines) > 1:
                info['address'] = " ".join(remaining_lines[1:])
        else:
            info['name'] = "Unknown"
            info['address'] = ""

        # Store raw text for manual verification if needed
        info['raw_extraction'] = raw_text.replace('\n', ' | ')
        
        return info

    def save_data(self):
        """Saves the list of dictionaries to CSV and Excel with deduplication."""
        df = pd.DataFrame(self.data_list)
        
        # --- ÉTAPE DE DÉDUPLICATION ---
        # 1. On liste toutes les colonnes à vérifier
        # On prend tout le monde SAUF 'source_file' et 'raw_extraction'
        cols_to_check = [c for c in df.columns if c not in ['source_file', 'raw_extraction']]
        
        # 2. On compte combien on en avait avant pour le log
        initial_count = len(df)
        
        # 3. On supprime les doublons
        # keep='first' : on garde la première occurrence trouvée
        # inplace=True : on modifie le dataframe directement
        df.drop_duplicates(subset=cols_to_check, keep='first', inplace=True)
        
        # 4. Petit log informatif
        removed_count = initial_count - len(df)
        if removed_count > 0:
            logger.info(f"Doublons supprimés : {removed_count} entrée(s).")
        # ------------------------------

        # Order columns nicely (Name, Phones, Emails, Address, Source)
        cols = list(df.columns)
        priority_cols = ['name', 'address']
        # Sort phones and emails dynamically
        phones = sorted([c for c in cols if c.startswith('phone')])
        emails = sorted([c for c in cols if c.startswith('email')])
        # On s'assure de garder les autres colonnes (dont source_file)
        others = [c for c in cols if c not in priority_cols + phones + emails]
        
        final_order = priority_cols + phones + emails + others
        # On réordonne, mais attention si une colonne n'existe plus (cas rare), on intersecte
        final_order = [c for c in final_order if c in df.columns]
        
        df = df[final_order]

        # Export
        try:
            df.to_excel("contacts_export.xlsx", index=False)
            df.to_csv("contacts_export.csv", index=False)
            logger.info("Successfully exported to 'contacts_export.xlsx' and 'contacts_export.csv'")
            
            print("\n" + "="*30)
            
            print("Extraction finished !")
            print(f"Total contacts found (unique) : {len(df)}")
            if removed_count > 0:
                print(f"Duplicates removed : {removed_count}")
                
            if self.failed_count > 0:
                print(f"Files failed/empty : {self.failed_count}")
                
            elapsed_time = time.time() - self.start_time
            print(f"Execution time : {elapsed_time:.2f} seconds")
            
            print("="*30)
        except Exception as e:
            logger.error(f"Error saving files: {e}")

    def close_word(self):
        if self.word_app:
            try:
                self.word_app.Quit()
            except:
                pass
            
if __name__ == "__main__":
    # Uses the directory where the script is located
    current_directory = os.getcwd()
    
    extractor = ContactExtractor(current_directory)
    extractor.run()