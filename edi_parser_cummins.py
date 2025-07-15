import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
from datetime import datetime
import os

class EDIDelforCumminsParser:
    def __init__(self, filepath=None):
        self.root = tk.Tk()
        self.root.title("EDI Cummins Parser")
        self.root.geometry("1200x800")
        self.header_info = {}
        self.partner_info = {}
        self.delivery_schedules = []
        self.line_items = []
        self.main_window = None
        self.setup_ui()
        if filepath:
            self.load_file(filepath)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="Načíst EDI soubor", command=self.load_file).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Zpět na hlavní okno", command=self.back_to_main).pack(side=tk.LEFT, padx=(10, 0))
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.info_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.info_frame, text="Základní informace")
        self.delivery_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.delivery_frame, text="Plán dodávek")
        self.stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_frame, text="Statistiky")
        self.setup_info_tab()
        self.setup_delivery_tab()
        self.setup_stats_tab()

    def setup_info_tab(self):
        text_frame = ttk.Frame(self.info_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.info_text = tk.Text(text_frame, wrap=tk.WORD, font=('Courier', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        self.info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_delivery_tab(self):
        tree_frame = ttk.Frame(self.delivery_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        columns = ('Položka', 'Popis', 'Datum od', 'Datum do', 'Množství', 'Jednotka', 'Typ', 'SCC')
        self.delivery_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.delivery_tree.heading(col, text=col)
            self.delivery_tree.column(col, width=120)
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.delivery_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.delivery_tree.xview)
        self.delivery_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.delivery_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_stats_tab(self):
        stats_frame = ttk.Frame(self.stats_frame)
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.stats_text = tk.Text(stats_frame, wrap=tk.WORD, font=('Courier', 10))
        stats_scrollbar = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=stats_scrollbar.set)
        self.stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def parse_date(self, date_str, format_code):
        try:
            if format_code == '102':
                return datetime.strptime(date_str, '%Y%m%d').strftime('%d.%m.%Y')
            else:
                return date_str
        except:
            return date_str

    def parse_edi_datetime(self, datetime_str):
        try:
            if ':' in datetime_str:
                date_part, time_part = datetime_str.split(':')
                full_date = '20' + date_part
                formatted_date = datetime.strptime(full_date, '%Y%m%d').strftime('%d.%m.%Y')
                formatted_time = datetime.strptime(time_part, '%H%M').strftime('%H:%M')
                return f"{formatted_date} {formatted_time}"
            return datetime_str
        except:
            return datetime_str

    def parse_edi_file(self, content):
        lines = content.strip().split("'")
        self.header_info = {}
        self.partner_info = {}
        self.delivery_schedules = []
        self.line_items = []
        
        current_item = {}
        last_lin_item = {}
        last_imd_desc = ''
        current_scc = ''
        deliveries_buffer = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            if line.startswith('UNB'):
                parts = line.split('+')
                if len(parts) >= 5:
                    self.header_info['Odesílatel'] = parts[2]
                    self.header_info['Příjemce_kód'] = parts[3]
                    self.header_info['Datum/Čas'] = self.parse_edi_datetime(parts[4])

            elif line.startswith('UNH'):
                parts = line.split('+')
                if len(parts) >= 2:
                    self.header_info['ID zprávy'] = parts[1]

            elif line.startswith('BGM'):
                parts = line.split('+')
                if len(parts) >= 3:
                    self.header_info['Číslo zprávy'] = parts[2]

            elif line.startswith('DTM'):
                parts = line.split('+')
                if len(parts) >= 2:
                    dtm_parts = parts[1].split(':')
                    if len(dtm_parts) >= 3:
                        code = dtm_parts[0]
                        value = dtm_parts[1]
                        fmt = dtm_parts[2]
                        formatted_date = self.parse_date(value, fmt)
                        if code == '137':
                            self.header_info['Datum dokumentu'] = formatted_date
                        elif code == '50':
                            current_item['Datum do'] = formatted_date
                        elif code == '2':
                            for d in deliveries_buffer:
                                d['Datum od'] = formatted_date
                                self.delivery_schedules.append(d)
                            deliveries_buffer.clear()

            elif line.startswith('NAD'):
                parts = line.split('+')
                if len(parts) >= 3:
                    role = parts[1]
                    if role in ['SU', 'ST']:
                        name_parts = [p.replace('?+', '').replace('?', '').strip() for p in parts[4:] if p]
                        if role == 'SU':
                            self.partner_info['Dodavatel'] = ' '.join(name_parts)
                        elif role == 'ST':
                            self.partner_info['Příjemce'] = ', '.join(name_parts)

            elif line.startswith('LIN'):
                if current_item:
                    self.line_items.append(current_item.copy())
                current_item = {}
                parts = line.split('+')
                if len(parts) >= 4:
                    current_item['Položka'] = parts[3].split(':')[0]
                last_lin_item = current_item.copy()
                current_scc = ''

            elif line.startswith('IMD'):
                parts = line.split('+')
                if len(parts) >= 5:
                    last_imd_desc = parts[4].replace(':', '').strip()
                    current_item['Popis'] = last_imd_desc

            elif line.startswith('SCC'):
                parts = line.split('+')
                if len(parts) >= 2:
                    current_scc = parts[1]

            elif line.startswith('QTY'):
                parts = line.split('+')
                if len(parts) >= 2:
                    qty_parts = parts[1].split(':')
                    if len(qty_parts) >= 3:
                        qty_type = qty_parts[0]
                        quantity = qty_parts[1]
                        unit = qty_parts[2]
                        typ = 'Neznámý'
                        if qty_type == '1':
                            typ = 'Dodávka'
                        elif qty_type == '3':
                            typ = 'Kumulativní'
                        elif qty_type == '48':
                            typ = 'Plánované'
                        delivery = {
                            'Položka': last_lin_item.get('Položka', ''),
                            'Popis': last_imd_desc,
                            'Množství': quantity,
                            'Jednotka': unit,
                            'Typ': typ,
                            'Datum do': last_lin_item.get('Datum do', ''),
                            'SCC': current_scc
                        }
                        deliveries_buffer.append(delivery)

        for d in deliveries_buffer:
            d['Datum od'] = ''
            self.delivery_schedules.append(d)

        if current_item:
            self.line_items.append(current_item.copy())

    def load_file(self, filepath=None):
        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.parse_edi_file(content)
                self.display_data()
            except Exception as e:
                messagebox.showerror("Chyba", f"Nelze načíst soubor: {str(e)}")

    def display_data(self):
        self.info_text.delete(1.0, tk.END)
        info_content = "=== HLAVIČKA DOKUMENTU ===\n"
        for key, value in self.header_info.items():
            if key != 'Příjemce_kód':
                info_content += f"{key}: {value}\n"
        info_content += "\n=== INFORMACE O PARTNERECH ===\n"
        for key, value in self.partner_info.items():
            info_content += f"{key}: {value}\n"
        self.info_text.insert(1.0, info_content)
        for item in self.delivery_tree.get_children():
            self.delivery_tree.delete(item)
        def datum_od_key(delivery):
            date_str = delivery.get('Datum od', '')
            try:
                return datetime.strptime(date_str, '%d.%m.%Y') if date_str else datetime.max
            except:
                return datetime.max
        sorted_deliveries = sorted(self.delivery_schedules, key=datum_od_key)
        for delivery in sorted_deliveries:
            self.delivery_tree.insert('', tk.END, values=(
                delivery.get('Položka', ''),
                delivery.get('Popis', ''),
                delivery.get('Datum od', ''),
                delivery.get('Datum do', ''),
                delivery.get('Množství', ''),
                delivery.get('Jednotka', ''),
                delivery.get('Typ', ''),
                delivery.get('SCC', '')
            ))
        self.stats_text.delete(1.0, tk.END)
        stats_content = "=== STATISTIKY ===\n"
        stats_content += f"Celkový počet dodávek: {len(self.delivery_schedules)}\n"
        total_qty = sum(int(d.get('Množství', 0)) for d in self.delivery_schedules if d.get('Množství', '').isdigit())
        stats_content += f"Celkové množství: {total_qty:,} kusů\n"
        self.stats_text.insert(1.0, stats_content)

    def back_to_main(self):
        """Closes the current window and returns to the main application"""
        # Close current window
        self.root.destroy()
        # Return to main window
        if self.main_window:
            self.main_window.root.deiconify()  # Show the main window if it exists
        else:
            # Create new main window instance
            app = EDIUnifiedParser()
            app.root.mainloop()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = EDIDelforCumminsParser()
    app.run()