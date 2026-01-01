"""
Agent Performance Data Processor - Native Windows GUI
Offline executable version with all functionality
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import io
import os
from pathlib import Path
import threading
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

class AgentPerformanceGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Agent Performance Data Processor")
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)
        
        # Variables
        self.df = None
        self.metadata_rows = []
        self.processed_df = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="📊 Agent Performance Data Processor", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # File selection
        ttk.Label(file_frame, text="CSV File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.file_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_var, state="readonly")
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.grid(row=0, column=2)
        
        process_btn = ttk.Button(file_frame, text="Process Data", command=self.process_data)
        process_btn.grid(row=0, column=3, padx=(10, 0))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Data tab
        self.data_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_frame, text="Processed Data")
        
        # Create treeview for data display
        self.setup_data_view()
        
        # Summary tab
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Summary")
        self.setup_summary_view()
        
        # Log tab
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="Log")
        self.setup_log_view()
        
        # Export frame
        export_frame = ttk.LabelFrame(main_frame, text="Export Options", padding="10")
        export_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        export_csv_btn = ttk.Button(export_frame, text="Export CSV", command=self.export_csv)
        export_csv_btn.grid(row=0, column=0, padx=(0, 10))
        
        export_excel_btn = ttk.Button(export_frame, text="Export Styled Excel", command=self.export_excel)
        export_excel_btn.grid(row=0, column=1)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select a CSV file to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def setup_data_view(self):
        """Setup the data view tab"""
        # Frame for treeview and scrollbars
        tree_frame = ttk.Frame(self.data_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview
        self.tree = ttk.Treeview(tree_frame)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
    def setup_summary_view(self):
        """Setup the summary view tab"""
        summary_text = scrolledtext.ScrolledText(self.summary_frame, wrap=tk.WORD)
        summary_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.summary_text = summary_text
        
    def setup_log_view(self):
        """Setup the log view tab"""
        log_text = scrolledtext.ScrolledText(self.log_frame, wrap=tk.WORD)
        log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.log_text = log_text
        
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def browse_file(self):
        """Browse for CSV file"""
        filename = filedialog.askopenfilename(
            title="Select Agent Performance CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.file_var.set(filename)
            self.log(f"Selected file: {filename}")
            
    def process_data(self):
        """Process the selected CSV file"""
        if not self.file_var.get():
            messagebox.showerror("Error", "Please select a CSV file first")
            return
            
        # Run processing in thread to prevent UI freezing
        threading.Thread(target=self._process_data_thread, daemon=True).start()
        
    def _process_data_thread(self):
        """Process data in background thread"""
        try:
            self.status_var.set("Processing data...")
            self.log("Starting data processing...")
            
            # Load and clean data
            self.df, self.metadata_rows = self.load_and_clean_data(self.file_var.get())
            if self.df is None:
                return
                
            # Process time columns
            self.df = self.process_time_columns(self.df)
            
            # Reorder and sort
            self.processed_df = self.reorder_and_sort(self.df)
            
            # Update UI in main thread
            self.root.after(0, self.update_ui_after_processing)
            
        except Exception as e:
            self.root.after(0, lambda: self.show_error(f"Error processing data: {str(e)}"))
            
    def update_ui_after_processing(self):
        """Update UI after data processing is complete"""
        try:
            # Update treeview
            self.update_treeview()
            
            # Update summary
            self.update_summary()
            
            # Switch to data tab
            self.notebook.select(0)
            
            self.status_var.set("Data processed successfully")
            self.log("Data processing completed successfully")
            
        except Exception as e:
            self.show_error(f"Error updating UI: {str(e)}")
            
    def load_and_clean_data(self, file_path):
        """Load CSV and perform initial cleaning"""
        try:
            self.log("Loading CSV file...")
            
            # Read file content
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            
            lines = content.split('\n')
            
            # Find the row that contains 'USER NAME' (the actual header)
            header_row = 0
            for i, line in enumerate(lines):
                if 'USER NAME' in line.upper():
                    header_row = i
                    break
            
            # Store ALL rows before the header as metadata
            metadata_rows = [line + '\n' for line in lines[:header_row] if line.strip()]
            
            # Create StringIO for pandas to read
            data_content = '\n'.join(lines[header_row:])
            df = pd.read_csv(io.StringIO(data_content), on_bad_lines='skip', engine='python')
            
            self.log(f"Loaded {len(df)} rows of data")
            
            # Columns to remove
            columns_to_delete = [
                'CURRENT USER GROUP', 'MOST RECENT USER GROUP', 'PAUSAVG', 'WAITAVG',
                'TALKAVG', 'DISPAVG', 'DEADAVG', 'CUSTAVG', 'ANS', 'SSMS', 'REDIAL',
                'test', 'testne', 'TestIT', 'TESTNC', 'TESTCB', 'Test22', 'DUPLICATE CALLS'
            ]
            
            # Drop columns (ignore if they don't exist)
            df = df.drop(columns=[col for col in columns_to_delete if col in df.columns], errors='ignore')
            
            # Remove last row (typically totals/summary)
            if len(df) > 0:
                df = df.iloc[:-1]
            
            self.log("Data cleaning completed")
            return df, metadata_rows
            
        except Exception as e:
            self.log(f"Error loading file: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error loading file: {str(e)}"))
            return None, None
            
    def process_time_columns(self, df):
        """Calculate total pause time from PAUSE, DEAD, and DISPO columns"""
        try:
            self.log("Processing time columns...")
            
            # Convert time columns to timedelta
            df['TOTAL PAUSE'] = (
                pd.to_timedelta(df['PAUSE'], errors='coerce').fillna(pd.Timedelta(0)) +
                pd.to_timedelta(df['DEAD'], errors='coerce').fillna(pd.Timedelta(0)) +
                pd.to_timedelta(df['DISPO'], errors='coerce').fillna(pd.Timedelta(0))
            )
            
            # Format as HH:MM:SS
            df['TOTAL PAUSE'] = df['TOTAL PAUSE'].apply(
                lambda x: f"{int(x.total_seconds() // 3600):02d}:"
                          f"{int((x.total_seconds() % 3600) // 60):02d}:"
                          f"{int(x.total_seconds() % 60):02d}"
                if pd.notna(x) else "00:00:00"
            )
            
            return df
            
        except Exception as e:
            self.log(f"Warning processing time columns: {str(e)}")
            return df
            
    def reorder_and_sort(self, df):
        """Reorder columns and sort by total inbound calls"""
        try:
            self.log("Reordering and sorting data...")
            
            # Convert ID to integer
            df['ID'] = pd.to_numeric(df['ID'], errors='coerce').fillna(0).astype(int)
            
            # Sort by total inbound calls (descending)
            if 'TOTAL INBOUND CALLS' in df.columns:
                df = df.sort_values(by='TOTAL INBOUND CALLS', ascending=False)
            
            # Reset index starting from 1
            df = df.reset_index(drop=True)
            df.index = df.index + 1
            
            # Reorder columns with ID first
            desired_columns = [
                'ID', 'USER NAME', 'CALLS', 'TIME', 'PAUSE', 'WAIT', 'TALK',
                'DISPO', 'DEAD', 'TOTAL PAUSE', 'CUSTOMER',
                'TOTAL INBOUND CALLS', 'TOTAL OUTBOUND CALLS'
            ]
            
            # Only include columns that exist
            existing_cols = [col for col in desired_columns if col in df.columns]
            df = df[existing_cols]
            
            # Add Remarks column as the last column
            df['REMARKS'] = ''
            
            # Add 'HD' in REMARKS if login hour (TIME) is less than 7 hours
            if 'TIME' in df.columns:
                for idx in df.index:
                    try:
                        time_val = pd.to_timedelta(df.loc[idx, 'TIME'])
                        if time_val < pd.to_timedelta('7:00:00'):
                            df.loc[idx, 'REMARKS'] = 'HD'
                    except:
                        pass
            
            return df
            
        except Exception as e:
            self.log(f"Warning reordering columns: {str(e)}")
            return df
            
    def update_treeview(self):
        """Update the treeview with processed data"""
        if self.processed_df is None:
            return
            
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Configure columns
        columns = list(self.processed_df.columns)
        self.tree['columns'] = columns
        self.tree['show'] = 'headings'
        
        # Configure column headings and widths
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=50)
            
        # Insert data
        for index, row in self.processed_df.iterrows():
            values = [str(row[col]) for col in columns]
            item = self.tree.insert('', 'end', values=values)
            
            # Apply color coding based on TOTAL INBOUND CALLS
            if 'TOTAL INBOUND CALLS' in columns:
                try:
                    calls = float(row['TOTAL INBOUND CALLS'])
                    if calls >= 70:
                        self.tree.set(item, 'TOTAL INBOUND CALLS', f"🟢 {calls}")
                    elif calls >= 60:
                        self.tree.set(item, 'TOTAL INBOUND CALLS', f"🟠 {calls}")
                    elif calls >= 50:
                        self.tree.set(item, 'TOTAL INBOUND CALLS', f"🟡 {calls}")
                    else:
                        self.tree.set(item, 'TOTAL INBOUND CALLS', f"🔴 {calls}")
                except:
                    pass
                    
    def update_summary(self):
        """Update the summary tab"""
        if self.processed_df is None:
            return
            
        summary = []
        summary.append("📊 AGENT PERFORMANCE SUMMARY")
        summary.append("=" * 50)
        summary.append("")
        
        # Basic statistics
        total_agents = len(self.processed_df)
        total_inbound = int(self.processed_df['TOTAL INBOUND CALLS'].sum())
        avg_inbound = self.processed_df['TOTAL INBOUND CALLS'].mean()
        
        summary.append(f"📈 Total Agents: {total_agents}")
        summary.append(f"📞 Total Inbound Calls: {total_inbound:,}")
        summary.append(f"📊 Average Inbound Calls: {avg_inbound:.2f}")
        summary.append("")
        
        # Top performer
        if len(self.processed_df) > 0:
            top_performer = self.processed_df.iloc[0]
            summary.append(f"🏆 Top Performer: {top_performer['USER NAME']}")
            summary.append(f"   Calls: {int(top_performer['TOTAL INBOUND CALLS'])}")
            summary.append("")
        
        # Performance distribution
        summary.append("📊 PERFORMANCE DISTRIBUTION:")
        summary.append("-" * 30)
        
        excellent = len(self.processed_df[self.processed_df['TOTAL INBOUND CALLS'] >= 70])
        good = len(self.processed_df[(self.processed_df['TOTAL INBOUND CALLS'] >= 60) & 
                                   (self.processed_df['TOTAL INBOUND CALLS'] < 70)])
        average = len(self.processed_df[(self.processed_df['TOTAL INBOUND CALLS'] >= 50) & 
                                      (self.processed_df['TOTAL INBOUND CALLS'] < 60)])
        below_avg = len(self.processed_df[self.processed_df['TOTAL INBOUND CALLS'] < 50])
        
        summary.append(f"🟢 Excellent (≥70 calls): {excellent} agents")
        summary.append(f"🟠 Good (60-69 calls): {good} agents")
        summary.append(f"🟡 Average (50-59 calls): {average} agents")
        summary.append(f"🔴 Below Average (<50 calls): {below_avg} agents")
        summary.append("")
        
        # HD (Half Day) analysis
        hd_count = len(self.processed_df[self.processed_df['REMARKS'] == 'HD'])
        summary.append(f"🟡 Half Day (HD) Agents: {hd_count}")
        summary.append("")
        
        # Color legend
        summary.append("🎨 COLOR LEGEND:")
        summary.append("-" * 20)
        summary.append("🟢 Green: ≥70 calls (Excellent)")
        summary.append("🟠 Orange: 60-69 calls (Good)")
        summary.append("🟡 Yellow: 50-59 calls (Average)")
        summary.append("🔴 Red: <50 calls (Below Average)")
        summary.append("🟡 HD: Login time <7 hours")
        
        # Update summary text
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(1.0, '\n'.join(summary))
        
    def export_csv(self):
        """Export data to CSV"""
        if self.processed_df is None:
            messagebox.showerror("Error", "No data to export. Please process a file first.")
            return
            
        filename = filedialog.asksavename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save CSV File"
        )
        
        if filename:
            try:
                # Create CSV with metadata
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    # Write metadata
                    for row in self.metadata_rows:
                        f.write(row)
                    
                    # Write data
                    self.processed_df.to_csv(f, index=False)
                
                messagebox.showinfo("Success", f"Data exported to {filename}")
                self.log(f"Data exported to CSV: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error exporting CSV: {str(e)}")
                
    def export_excel(self):
        """Export data to styled Excel"""
        if self.processed_df is None:
            messagebox.showerror("Error", "No data to export. Please process a file first.")
            return
            
        filename = filedialog.asksavename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Excel File"
        )
        
        if filename:
            try:
                self.status_var.set("Creating Excel file...")
                self.log("Creating styled Excel file...")
                
                # Run export in thread
                threading.Thread(
                    target=self._export_excel_thread, 
                    args=(filename,), 
                    daemon=True
                ).start()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error exporting Excel: {str(e)}")
                
    def _export_excel_thread(self, filename):
        """Export Excel in background thread"""
        try:
            output = io.BytesIO()
            
            # Create initial Excel file
            num_metadata_rows = len(self.metadata_rows)
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                self.processed_df.to_excel(writer, index=False, startrow=num_metadata_rows, sheet_name='Agent Performance')
            
            # Load and modify workbook for styling
            output.seek(0)
            wb = load_workbook(output)
            ws = wb.active
            
            # Apply styling (simplified version)
            header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            header_font = Font(bold=True, color='000000')
            
            # Style header row
            header_row_num = num_metadata_rows + 1
            for col in range(1, len(self.processed_df.columns) + 1):
                cell = ws.cell(row=header_row_num, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Save file
            wb.save(filename)
            
            # Update UI in main thread
            self.root.after(0, lambda: self._excel_export_complete(filename))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error creating Excel file: {str(e)}"))
            
    def _excel_export_complete(self, filename):
        """Called when Excel export is complete"""
        messagebox.showinfo("Success", f"Styled Excel file created: {filename}")
        self.log(f"Styled Excel exported: {filename}")
        self.status_var.set("Excel export completed")
        
    def show_error(self, message):
        """Show error message"""
        messagebox.showerror("Error", message)
        self.status_var.set("Error occurred")
        self.log(f"ERROR: {message}")
        
    def run(self):
        """Run the application"""
        self.log("Agent Performance Data Processor started")
        self.log("Select a CSV file and click 'Process Data' to begin")
        self.root.mainloop()

if __name__ == "__main__":
    app = AgentPerformanceGUI()
    app.run()