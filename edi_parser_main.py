import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from edi_parser_cummins import EDIDelforCumminsParser
from edi_parser_trwkob import EDITrwkobParser
from edi_parser_minebea import EDIDelforParser as EDIDelforMinebeaParser

class EDIUnifiedParser:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("EDI Unified Parser")
        self.root.geometry("600x400")
        self.setup_ui()
        # Store reference to main window instance
        self.main_window = self

    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(btn_frame, text="Načíst EDI soubor", command=self.load_file).pack(side=tk.LEFT)
        
        self.info_text = tk.Text(main_frame, wrap=tk.WORD, font=('Courier', 10))
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        self.info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Vyberte EDI soubor",
            filetypes=[("EDI files", "*.edi"), ("All files", "*.*")]
        )
        
        if not filepath:
            return

        try:
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            
            # Detect file type based on both filename and content
            file_type = self.detect_file_type(filepath, content)
            
            # Function to run parser and close main window
            def run_parser(parser_func):
                success = parser_func(filepath)
                if success:
                    if file_type != "trwkob":
                        self.root.destroy()  # Close main window only for non-Trwkob parsers
                return success
            
            if file_type == "cummins":
                success = run_parser(self.run_cummins_parser)
            elif file_type == "trwkob":
                success = run_parser(self.run_trwkob_parser)
            elif file_type == "minebea":
                success = run_parser(self.run_minebea_parser)
            else:
                messagebox.showerror("Chyba", "Nepodporovaný typ souboru")
                
        except Exception as e:
            messagebox.showerror("Chyba", f"Chyba při načítání souboru: {str(e)}")
            # Keep the main window open on error
            return False

    def detect_file_type(self, filepath, content):
        # Look for patterns in both filename and content
        filename = os.path.basename(filepath).upper()
        content_upper = content.upper()
        
        # Check for Cummins patterns (both in filename and content)
        cummins_patterns = [
            "CUMMINS", "CMI", "CMI-", "CMI_",
            "DELFOR_CUMMINS", "CUMMINS_DELFOR"
        ]
        if any(pattern in filename for pattern in cummins_patterns) or \
           any(pattern in content_upper for pattern in cummins_patterns):
            return "cummins"
            
        # Check for Minebea patterns
        minebea_patterns = [
            "MINEBEA", "MINOL", "MINEBEA-MINOL", "MBM",
            "DELFOR_MINEBEA", "MINEBEA_DELFOR"
        ]
        if any(pattern in filename for pattern in minebea_patterns) or \
           any(pattern in content_upper for pattern in minebea_patterns):
            return "minebea"
            
        # Check for Trwkob patterns
        trwkob_patterns = [
            "TRWKOB", "TRW-KOB", "TRW_KOB", "KOBALT",
            "DELFOR_TRWKOB", "TRWKOB_DELFOR"
        ]
        if any(pattern in filename for pattern in trwkob_patterns) or \
           any(pattern in content_upper for pattern in trwkob_patterns):
            return "trwkob"
            
        # If no specific pattern found, try to detect by file structure
        if content.startswith("UNB") or content.startswith("UNA"):
            # This is a standard EDI file structure
            return "minebea"  # Default to Minebea as fallback
            
        return None

    def run_cummins_parser(self, filepath):
        # Create parser instance but don't start mainloop yet
        parser = EDIDelforCumminsParser(filepath)
        parser.main_window = self.main_window
        success = parser.load_file(filepath)
        if success:
            # Close main window
            self.root.destroy()
            # Start parser's mainloop
            parser.root.mainloop()
        return success

    def run_trwkob_parser(self, filepath):
        # Create parser instance but don't start mainloop yet
        parser = EDITrwkobParser(filepath)
        parser.main_window = self.main_window
        try:
            success = parser.load_file(filepath)
            if success:
                # Start parser's mainloop without closing main window
                parser.root.mainloop()
            return success
        except Exception as e:
            messagebox.showerror("Chyba", f"Chyba při načítání souboru: {str(e)}")
            return False

    def run_minebea_parser(self, filepath):
        # Create parser instance but don't start mainloop yet
        parser = EDIDelforMinebeaParser(filepath)
        parser.main_window = self.main_window
        success = parser.load_file(filepath)
        if success:
            # Close main window
            self.root.destroy()
            # Start parser's mainloop
            parser.root.mainloop()
        return success

def main():
    app = EDIUnifiedParser()
    app.root.mainloop()

if __name__ == "__main__":
    main()
