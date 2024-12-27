import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import yfinance as yf
import threading
import time
import logging

class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Yahoo Finance Data Puller")
        
        self.default_columns = ['Industry', 'Sector', 'Market Cap']
        self.columns = ['Ticker'] + self.default_columns
        self.selected_columns = self.default_columns.copy()
        
        # Set up logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        
        # Create a frame for the input box
        self.input_frame = ttk.Frame(root)
        self.input_frame.pack(fill='x', padx=10, pady=10)

        # Create input text box for tickers
        self.input_text = tk.Text(self.input_frame, height=2)
        self.input_text.pack(fill='x')

        # Create a frame for the table
        self.frame = ttk.Frame(root)
        self.frame.pack(fill='both', expand=True)
        
        # Create a Treeview table for tickers and data
        self.table = ttk.Treeview(self.frame, columns=self.columns, show='headings')
        self.setup_table()
        self.table.pack(side='left', fill='both', expand=True)
        
        # Insert initial empty rows
        for _ in range(10):
            self.table.insert('', 'end', values=['' for _ in self.columns])
        
        # Create buttons
        self.load_button = ttk.Button(root, text='Load Tickers', command=self.load_tickers, underline=5)
        self.load_button.pack(side='left', padx=(15, 5))

        self.select_columns_button = ttk.Button(root, text='Select Columns', command=self.select_columns, underline=7)
        self.select_columns_button.pack(side='left', padx=5)

        self.fetch_button = ttk.Button(root, text='Fetch Data', command=self.handle_data_pull, underline=6)
        self.fetch_button.pack(side='left', padx=(5, 15))

        self.export_button = ttk.Button(root, text='Export to CSV', command=self.export_to_csv, underline=0)
        self.export_button.pack(side='right', padx=15)
        
        # Create a loading bar
        self.progress = ttk.Progressbar(root, orient='horizontal', length=200, mode='determinate')
        self.progress.pack(side='right', padx=15, pady=10)
        
        # Create progress label
        self.progress_label = ttk.Label(root)
        self.progress_label.pack(side='right', padx=15, pady=10)

        # Bind hotkeys
        root.bind('<Alt-t>', lambda event: self.load_tickers())
        root.bind('<Alt-c>', lambda event: self.select_columns())
        root.bind('<Alt-d>', lambda event: self.handle_data_pull())
        root.bind('<Alt-e>', lambda event: self.export_to_csv())

    def setup_table(self):
        self.table['columns'] = ['Ticker'] + self.selected_columns
        for col in self.table['columns']:
            self.table.heading(col, text=col)
            self.table.column(col, width=150, anchor='center' if col == 'Market Cap' else 'w')

    def load_tickers(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
            if file_path:
                df = pd.read_csv(file_path)
                tickers = df.iloc[:, 0].tolist()
                self.input_text.delete(1.0, tk.END)
                self.input_text.insert(tk.END, ', '.join(tickers[:10]))
                logging.info("Tickers loaded successfully from CSV file.")
        except Exception as e:
            logging.error(f"Failed to load tickers: {e}")
            messagebox.showerror("Error", "Failed to load tickers.")

    def fetch_data(self, ticker):
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            market_cap_inr = info.get('marketCap', 0) / 10**7  # Convert to crores INR
            row = [ticker.upper()] + [info.get(col.lower(), '') for col in self.selected_columns]
            if 'Market Cap' in self.selected_columns:
                row[self.selected_columns.index('Market Cap') + 1] = f"{market_cap_inr:,.2f} Cr"
            logging.info(f"Data fetched successfully for ticker: {ticker}")
        except Exception as e:
            logging.error(f"Failed to fetch data for ticker {ticker}: {e}")
            row = [ticker.upper()] + ['' for _ in self.selected_columns]
        return row

    def export_to_csv(self):
        try:
            data = []
            for child in self.table.get_children():
                data.append(self.table.item(child)['values'])
            df = pd.DataFrame(data, columns=self.table['columns'])
            file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')])
            if file_path:
                df.to_csv(file_path, index=False)
                messagebox.showinfo('Success', 'Data exported successfully!')
                logging.info("Data exported successfully to CSV.")
        except Exception as e:
            logging.error(f"Failed to export data: {e}")
            messagebox.showerror("Error", "Failed to export data.")

    def display_data(self, data):
        for i, row in enumerate(data):
            self.table.item(self.table.get_children()[i], values=row)

    def handle_data_pull(self):
        tickers = self.input_text.get(1.0, tk.END).replace(" ", "").strip().split(',')
        tickers = [ticker.strip().upper() for ticker in tickers if ticker.strip()]
        if tickers:
            self.progress['value'] = 0
            self.progress_label.config(text="")
            self.progress.start()
            threading.Thread(target=self.fetch_and_display_data, args=(tickers,)).start()

    def fetch_and_display_data(self, tickers):
        total_tickers = len(tickers)
        data = []
        for i, ticker in enumerate(tickers):
            row = self.fetch_data(ticker)
            data.append(row)
            self.progress['value'] = (i + 1) / total_tickers * 100
            time.sleep(0.5)  # Simulate data fetching delay for smoother UI experience
        self.progress.stop()
        self.progress['value'] = 100
        self.progress_label.config(text="Data fetch complete")
        self.display_data(data)

    def select_columns(self):
        self.column_window = tk.Toplevel(self.root)
        self.column_window.title("Select Columns")
        
        self.column_frame = ttk.Frame(self.column_window)
        self.column_frame.pack(padx=10, pady=10, fill='both', expand=True)

        self.column_checkboxes = []

        # Example list of columns (can be updated with all available columns from Yahoo Finance)
        col_values = [
            'Industry', 'Sector', 'Market Cap', 
            'Volume', 'Open', 'Close', 'High', 'Low',
            'EBITDA', 'EPS', 'PE Ratio', 'Dividend Yield'
        ]
        
        for i, col in enumerate(col_values):
            var = tk.BooleanVar(value=(col in self.selected_columns))
            chk = ttk.Checkbutton(self.column_frame, text=col, variable=var)
            chk.grid(row=i // 3, column=i % 3, sticky='w', padx=5, pady=5)
            self.column_checkboxes.append((col, var))

        self.ok_button = ttk.Button(self.column_window, text='OK', command=self.apply_column_selection)
        self.ok_button.pack(side='left', padx=10, pady=10)

        self.cancel_button = ttk.Button(self.column_window, text='Cancel', command=self.column_window.destroy)
        self.cancel_button.pack(side='right', padx=10, pady=10)

    def apply_column_selection(self):
        self.selected_columns = [col for col, var in self.column_checkboxes if var.get()]
        self.columns = ['Ticker'] + self.selected_columns
        self.setup_table()
        self.column_window.destroy()

if __name__ == "__main__":
    # Check if an instance is already running, if so, update it, else start a new instance.
    if 'app' in globals():
        app.handle_data_pull()
    else:
        root = tk.Tk()
        app = FinanceApp(root)
        root.mainloop()
