import os
import re
import pandas as pd
import pdfplumber
import docx
from odf import text, teletype
from odf.opendocument import load
from pathlib import Path
import logging

# --- Configuration ---
# Configure logging to see what happens in the console
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

class ContactExtractor:
    def __init__(self, root_dir):
        self.root_dir = root_dir
        self.data_list = []
        
        # Regex for French phone numbers (flexible: 06, +33, spaces, dots)
        self.phone_pattern = re.compile(r'(?:(?:\+|00)33|0)\s*[1-9](?:[\s.-]*\d{2}){4}')
        # Regex for Emails
        self.email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

    def run(self):
        """Main execution method."""
        logger.info(f"Starting scan in: {self.root_dir}")
        
        # Walk through directory recursively
        for root, dirs, files in os.walk(self.root_dir):
            for file in files:
                file_path = os.path.join(root, file)
                self.process_file(file_path)

        if not self.data_list:
            logger.warning("No data extracted. Check your files or table structures.")
            return

        self.save_data()

    def process_file(self, file_path):
        """Dispatcher that calls the right extractor based on extension."""
        ext = os.path.splitext(file_path)[1].lower()
        raw_text = None

        try:
            if ext == '.pdf':
                raw_text = self.extract_from_pdf(file_path)
            elif ext == '.docx':
                raw_text = self.extract_from_docx(file_path)
            elif ext == '.odt':
                raw_text = self.extract_from_odt(file_path)
            elif ext == '.doc':
                logger.warning(f"Skipping .doc file (requires conversion to .docx): {file_path}")
                return
            else:
                return # Ignore other files
            
            if raw_text:
                logger.info(f"Data found in: {os.path.basename(file_path)}")
                parsed_info = self.parse_contact_info(raw_text)
                parsed_info['source_file'] = file_path # Add source path for reference
                self.data_list.append(parsed_info)
            else:
                logger.debug(f"No valid table data found in {os.path.basename(file_path)}")

        except Exception as e:
            logger.error(f"Error processing {file_path}: {e}")

    def extract_from_pdf(self, file_path):
        """Extracts text from the 2nd cell of the 1st table in a PDF."""
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return None
            first_page = pdf.pages[0]
            tables = first_page.extract_tables()
            
            if tables:
                first_table = tables[0]
                # Flatten the table to find the 2nd cell regardless of row/col structure
                # We assume reading order: Row 1 Col 1, Row 1 Col 2, etc.
                flattened_cells = [cell for row in first_table for cell in row if cell]
                
                if len(flattened_cells) >= 2:
                    return flattened_cells[1] # Return the second cell
        return None

    def extract_from_docx(self, file_path):
        """Extracts text from the 2nd cell of the 1st table in a DOCX."""
        doc = docx.Document(file_path)
        if doc.tables:
            table = doc.tables[0]
            # Strategy: Get all cells in the first row, then second row if needed
            # We want the second "physical" cell. 
            # Usually Row 0, Col 1.
            
            all_cells = []
            for row in table.rows:
                for cell in row.cells:
                    all_cells.append(cell.text.strip())
                    if len(all_cells) >= 2:
                        break
                if len(all_cells) >= 2:
                    break
            
            if len(all_cells) >= 2:
                return all_cells[1]
        return None

    def extract_from_odt(self, file_path):
        """Extracts text from the 2nd cell of the 1st table in an ODT."""
        doc = load(file_path)
        tables = doc.getElementsByType(text.Table)
        
        if tables:
            table = tables[0]
            rows = table.getElementsByType(text.TableRow)
            
            all_cells_text = []
            for row in rows:
                cells = row.getElementsByType(text.TableCell)
                for cell in cells:
                    all_cells_text.append(teletype.extractText(cell).strip())
                    if len(all_cells_text) >= 2:
                        break
                if len(all_cells_text) >= 2:
                    break
            
            if len(all_cells_text) >= 2:
                return all_cells_text[1]
        return None

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

        # 4. Heuristic for Name and Address
        # We assume lines that are NOT phones or emails constitute the entity/address
        remaining_lines = []
        for line in lines:
            # Check if line is purely a phone number or email extracted above
            # This is a simple check; text extraction can be messy so we keep the line 
            # if it contains partial address info even if it has a phone.
            is_purely_contact = False
            for p in phones:
                if p in line and len(line) < len(p) + 5: is_purely_contact = True
            for e in emails:
                if e in line and len(line) < len(e) + 5: is_purely_contact = True
            
            if not is_purely_contact:
                remaining_lines.append(line)

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
        """Saves the list of dictionaries to CSV and Excel."""
        df = pd.DataFrame(self.data_list)
        
        # Order columns nicely (Name, Phones, Emails, Address, Source)
        cols = list(df.columns)
        priority_cols = ['name', 'address']
        # Sort phones and emails dynamically
        phones = sorted([c for c in cols if c.startswith('phone')])
        emails = sorted([c for c in cols if c.startswith('email')])
        others = [c for c in cols if c not in priority_cols + phones + emails]
        
        final_order = priority_cols + phones + emails + others
        df = df[final_order]

        # Export
        try:
            df.to_excel("contacts_export.xlsx", index=False)
            df.to_csv("contacts_export.csv", index=False)
            logger.info("Successfully exported to 'contacts_export.xlsx' and 'contacts_export.csv'")
            print("\n" + "="*30)
            print("Extraction finished !")
            print(f"Total contacts found : {len(df)}")
            print("="*30)
        except Exception as e:
            logger.error(f"Error saving files: {e}")

if __name__ == "__main__":
    # Uses the directory where the script is located
    current_directory = os.getcwd()
    
    extractor = ContactExtractor(current_directory)
    extractor.run()