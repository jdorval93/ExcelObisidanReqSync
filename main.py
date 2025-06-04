import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime
import re


# =============================================================================
# CONFIGURATION - Edit these settings as needed
# =============================================================================

# Default file paths (leave empty to use UI selection)
DEFAULT_EXCEL_FILE = "C:\\Users\\jdorval\\Desktop\\Requirements Traceability Matrix IFSCloud.xlsx" 
DEFAULT_OBSIDIAN_VAULT = "C:\\Obsidian\\Hecla\\Work knowledge\\1.Projects\\Project keystone\\Requirements"  

# Auto-generate overview file after creating files
AUTO_GENERATE_OVERVIEW = True

# =============================================================================

class RequirementsConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Requirements to Obsidian MD Converter")
        self.root.geometry("600x500")
        
        # Variables - Initialize with default values from config
        self.excel_file = tk.StringVar(value=DEFAULT_EXCEL_FILE)
        self.output_dir = tk.StringVar(value=DEFAULT_OBSIDIAN_VAULT)
        self.status_text = tk.StringVar(value="Ready to create files from Excel...")
        
        self.setup_gui()
        
    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel → Obsidian Requirements Creator", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # Subtitle
        subtitle = ttk.Label(main_frame, text="Creates missing requirement files - updates done manually", 
                            font=('Arial', 10), foreground="gray")
        subtitle.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Excel file selection
        ttk.Label(main_frame, text="Excel File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_file, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_excel_file).grid(row=2, column=2, padx=5)
        
        # Output directory selection
        ttk.Label(main_frame, text="Obsidian Vault:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=3, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_output_dir).grid(row=3, column=2, padx=5)
        
        # Column mapping info
        info_frame = ttk.LabelFrame(main_frame, text="Column Mapping", padding="10")
        info_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        mapping_text = """Column A: Requirement ID (REQUIRED for filename)
Column B: Category/Functional Activity  
Column C: Topic
Column E: Short Description (REQUIRED for filename)
Column F: Description Overview
Column G: Priority

Note: Both Column A and E are required to create files."""
        
        ttk.Label(info_frame, text=mapping_text, font=('Arial', 9), foreground="gray").grid(row=0, column=0, sticky=tk.W)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=15)
        
        # Check what's missing button
        ttk.Button(button_frame, text="Check Missing Files", 
                  command=self.check_missing_files,
                  style='TButton').grid(row=0, column=0, padx=5)
        
        # Create missing files button
        ttk.Button(button_frame, text="Create Missing Files", 
                  command=self.create_missing_files,
                  style='Accent.TButton').grid(row=0, column=1, padx=5)
        
        # Overview generation button
        ttk.Button(button_frame, text="Generate Overview", 
                  command=self.generate_overview_only,
                  style='TButton').grid(row=0, column=2, padx=5)
        
        # Status
        ttk.Label(main_frame, text="Status:").grid(row=6, column=0, sticky=tk.W, pady=(10, 0))
        status_label = ttk.Label(main_frame, textvariable=self.status_text, 
                                foreground="blue")
        status_label.grid(row=6, column=1, columnspan=2, sticky=tk.W, padx=5)
        
        # Log text area
        ttk.Label(main_frame, text="Log:").grid(row=7, column=0, sticky=(tk.W, tk.N), pady=(10, 0))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = tk.Text(log_frame, height=20, width=70)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(8, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log default settings if they're configured
        if DEFAULT_EXCEL_FILE:
            self.log(f"Using default Excel file: {DEFAULT_EXCEL_FILE}")
        if DEFAULT_OBSIDIAN_VAULT:
            self.log(f"Using default Obsidian vault: {DEFAULT_OBSIDIAN_VAULT}")
            
    def browse_excel_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            
    def browse_output_dir(self):
        dirname = filedialog.askdirectory(title="Select Obsidian Vault Directory")
        if dirname:
            self.output_dir.set(dirname)
            
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def sanitize_filename(self, text):
        """Convert text to a safe filename"""
        # Remove or replace invalid characters
        text = re.sub(r'[<>:"/\\|?*]', '_', text)
        text = re.sub(r'\s+', '_', text)  # Replace spaces with underscores
        text = text.strip('._')  # Remove leading/trailing dots and underscores
        return text[:100]  # Limit length
        
    def generate_filename(self, row):
        """Generate filename from Excel row data - requires both ID and short description"""
        req_id = str(row['A']).strip() if pd.notna(row['A']) and str(row['A']).strip() else ""
        short_desc = str(row['E']).strip() if pd.notna(row['E']) and str(row['E']).strip() else ""
        
        # Only create filename if both ID and short description are available
        if not req_id:
            raise ValueError("Missing Requirement ID (Column A)")
        
        if not short_desc:
            raise ValueError("Missing Short Description (Column E)")
        
        # Combine ID and short description for filename
        filename_base = f"{req_id}_{short_desc}"
        filename_base = self.sanitize_filename(filename_base)
        return f"{filename_base}.md"
        
    def create_md_content(self, row_data):
        """Create markdown content from row data"""
        content = []
        
        # Create title from Requirement ID and Short Description
        req_id = str(row_data['A']).strip() if pd.notna(row_data['A']) else "Unknown ID"
        short_desc = str(row_data['E']).strip() if pd.notna(row_data['E']) else "No Description"
        
        content.append(f"# {req_id} - {short_desc}")
        content.append("")
        
        # Create a clean table with the requirement data
        content.append("| Attribute | Value |")
        content.append("|-----------|-------|")
        
        # Map columns to meaningful names
        column_names = {
            'A': 'Requirement ID',
            'B': 'Category/Functional Activity',
            'C': 'Topic', 
            'E': 'Short Description',
            'F': 'Description Overview',
            'G': 'Priority'
        }
        
        columns = ['A', 'B', 'C', 'E', 'F', 'G']
        for col in columns:
            if col in row_data and pd.notna(row_data[col]):
                field_name = column_names.get(col, f'Column {col}')
                # Escape pipe characters in values and clean up text
                value = str(row_data[col]).replace('|', '\\|').replace('\n', ' ').strip()
                content.append(f"| **{field_name}** | {value} |")
        
        content.append("")
        content.append("---")
        content.append("")
        content.append("## Notes")
        content.append("*Add your additional notes and details here...*")
        content.append("")
        
        # Add creation metadata
        content.append("<!-- CREATION_METADATA")
        content.append(f"created_by: excel_to_obsidian_converter")
        content.append(f"creation_date: {datetime.now().isoformat()}")
        content.append("-->")
        content.append("")
        
        return "\n".join(content)
        
    def read_excel_requirements(self):
        """Read and parse Excel file requirements"""
        try:
            # Read Excel file
            df = pd.read_excel(self.excel_file.get(), header=None)
            
            # Extract data from row 6 onward (index 5), columns A, B, C, E, F, G
            columns = ['A', 'B', 'C', 'E', 'F', 'G']
            col_indices = [0, 1, 2, 4, 5, 6]  # A=0, B=1, C=2, E=4, F=5, G=6 (0-indexed)
            
            # Get data from row 6 onward
            data_rows = df.iloc[5:, col_indices].copy()
            data_rows.columns = columns
            
            # Remove completely empty rows
            data_rows = data_rows.dropna(how='all')
            
            # Filter to meaningful rows (those with key data)
            meaningful_rows = []
            for idx, row in data_rows.iterrows():
                if not (pd.isna(row['A']) and pd.isna(row['C']) and pd.isna(row['E']) and pd.isna(row['F'])):
                    meaningful_rows.append((idx, row))
            
            return meaningful_rows
            
        except Exception as e:
            raise Exception(f"Error reading Excel file: {e}")
    
    def check_missing_files(self):
        """Check which Excel requirements are missing files"""
        if not self.excel_file.get():
            messagebox.showerror("Error", "Please select an Excel file")
            return
            
        if not self.output_dir.get():
            messagebox.showerror("Error", "Please select output directory")
            return
            
        try:
            self.status_text.set("Checking missing files...")
            self.log("=" * 60)
            self.log("CHECKING FOR MISSING REQUIREMENT FILES")
            self.log("=" * 60)
            
            # Read Excel requirements
            requirements = self.read_excel_requirements()
            self.log(f"📊 Found {len(requirements)} meaningful requirements in Excel")
            
            # Check which files exist
            missing_files = []
            existing_files = []
            invalid_requirements = []
            
            for idx, row in requirements:
                try:
                    filename = self.generate_filename(row)
                    filepath = os.path.join(self.output_dir.get(), filename)
                    
                    req_id = str(row['A']).strip() if pd.notna(row['A']) else "No ID"
                    short_desc = str(row['E']).strip() if pd.notna(row['E']) else "No description"
                    
                    if os.path.exists(filepath):
                        existing_files.append((filename, req_id, short_desc))
                        self.log(f"✓ EXISTS: {filename}")
                    else:
                        missing_files.append((filename, req_id, short_desc, idx, row))
                        self.log(f"❌ MISSING: {filename} ({req_id} - {short_desc})")
                        
                except ValueError as e:
                    invalid_requirements.append((idx, row, str(e)))
                    row_num = idx + 6  # Adjust for Excel row numbering
                    req_id = str(row['A']).strip() if pd.notna(row['A']) and str(row['A']).strip() else "Missing"
                    short_desc = str(row['E']).strip() if pd.notna(row['E']) and str(row['E']).strip() else "Missing"
                    self.log(f"⚠ INVALID (Row {row_num}): {str(e)} - ID: '{req_id}', Short Desc: '{short_desc}'")
            
            # Summary
            self.log("=" * 60)
            self.log("SUMMARY")
            self.log(f"📊 Total requirements: {len(requirements)}")
            self.log(f"✓ Existing files: {len(existing_files)}")
            self.log(f"❌ Missing files: {len(missing_files)}")
            self.log(f"⚠ Invalid requirements: {len(invalid_requirements)}")
            
            if missing_files:
                self.log("=" * 60)
                self.log("MISSING REQUIREMENTS:")
                for filename, req_id, short_desc, idx, row in missing_files:
                    self.log(f"  • {req_id}: {short_desc}")
                self.log(f"\nUse 'Create Missing Files' to create these {len(missing_files)} files")
            
            if invalid_requirements:
                self.log("=" * 60)
                self.log("INVALID REQUIREMENTS (cannot create files):")
                for idx, row, error in invalid_requirements:
                    row_num = idx + 6
                    req_id = str(row['A']).strip() if pd.notna(row['A']) and str(row['A']).strip() else "Missing"
                    short_desc = str(row['E']).strip() if pd.notna(row['E']) and str(row['E']).strip() else "Missing"
                    self.log(f"  • Row {row_num}: {error} (ID: '{req_id}', Short Desc: '{short_desc}')")
            
            if not missing_files and not invalid_requirements:
                self.log("🎉 All valid requirements have corresponding files!")
            
            self.status_text.set(f"Check complete: {len(missing_files)} missing, {len(invalid_requirements)} invalid")
            
        except Exception as e:
            error_msg = f"Error checking files: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.status_text.set("Error occurred")
            messagebox.showerror("Error", error_msg)
    
    def create_missing_files(self):
        """Create files for Excel requirements that don't have corresponding files"""
        if not self.excel_file.get():
            messagebox.showerror("Error", "Please select an Excel file")
            return
            
        if not self.output_dir.get():
            messagebox.showerror("Error", "Please select output directory")
            return
            
        try:
            self.status_text.set("Creating missing files...")
            self.log("=" * 60)
            self.log("CREATING MISSING REQUIREMENT FILES")
            self.log("=" * 60)
            
            # Read Excel requirements
            requirements = self.read_excel_requirements()
            self.log(f"📊 Found {len(requirements)} meaningful requirements in Excel")
            
            # Create missing files
            created_count = 0
            skipped_count = 0
            error_count = 0
            
            # Ensure output directory exists
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            for idx, row in requirements:
                try:
                    filename = self.generate_filename(row)
                    filepath = os.path.join(self.output_dir.get(), filename)
                    
                    req_id = str(row['A']).strip() if pd.notna(row['A']) else "No ID"
                    short_desc = str(row['E']).strip() if pd.notna(row['E']) else "No description"
                    
                    if os.path.exists(filepath):
                        skipped_count += 1
                        self.log(f"⏭ SKIPPED: {filename} (already exists)")
                    else:
                        # Create the file
                        content = self.create_md_content(row)
                        
                        with open(filepath, 'w', encoding='utf-8') as f:
                            f.write(content)
                        
                        created_count += 1
                        self.log(f"✅ CREATED: {filename} ({req_id} - {short_desc})")
                        
                except ValueError as e:
                    error_count += 1
                    row_num = idx + 6  # Adjust for Excel row numbering (data starts at row 6)
                    req_id = str(row['A']).strip() if pd.notna(row['A']) and str(row['A']).strip() else "Missing"
                    short_desc = str(row['E']).strip() if pd.notna(row['E']) and str(row['E']).strip() else "Missing"
                    self.log(f"❌ ERROR (Row {row_num}): {str(e)} - ID: '{req_id}', Short Desc: '{short_desc}'")
                except Exception as e:
                    error_count += 1
                    row_num = idx + 6
                    self.log(f"❌ UNEXPECTED ERROR (Row {row_num}): {str(e)}")
            
            # Summary
            self.log("=" * 60)
            self.log("CREATION COMPLETE")
            self.log(f"✅ Files created: {created_count}")
            self.log(f"⏭ Files skipped: {skipped_count}")
            self.log(f"❌ Errors (files not created): {error_count}")
            self.log(f"📊 Total processed: {len(requirements)}")
            
            if error_count > 0:
                self.log(f"\n⚠ {error_count} requirements could not be processed due to missing ID or Short Description")
            
            self.status_text.set("File creation complete!")
            
            # Generate overview if enabled and files were created
            if AUTO_GENERATE_OVERVIEW and created_count > 0:
                self.log("Auto-generating requirements overview...")
                self.generate_overview_only()
            
            # Show summary dialog
            if error_count > 0:
                summary_msg = (f"File Creation Complete with Errors!\n\n"
                              f"Created: {created_count} new files\n"
                              f"Skipped: {skipped_count} existing files\n"
                              f"Errors: {error_count} requirements missing required data\n"
                              f"Total: {len(requirements)} requirements processed\n\n"
                              f"Check the log for details on failed requirements.\n"
                              f"Requirements need both ID (Column A) and Short Description (Column E).")
            else:
                summary_msg = (f"File Creation Complete!\n\n"
                              f"Created: {created_count} new files\n"
                              f"Skipped: {skipped_count} existing files\n"
                              f"Total: {len(requirements)} requirements processed\n\n"
                              f"All files are now in your Obsidian vault!")
                
            if error_count > 0:
                messagebox.showwarning("Creation Complete with Errors", summary_msg)
            else:
                messagebox.showinfo("Creation Complete", summary_msg)
            
        except Exception as e:
            error_msg = f"Error creating files: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.status_text.set("Error occurred")
            messagebox.showerror("Error", error_msg)
    
    def extract_requirement_data_from_file(self, filepath):
        """Extract requirement data from a markdown file"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Extract data from the table in the markdown
            req_data = {
                'filename': os.path.basename(filepath),
                'filepath': filepath,
                'requirement_id': '',
                'category': '',
                'topic': '',
                'short_description': '',
                'description': '',
                'priority': ''
            }
            
            # Parse the markdown table to extract requirement data
            lines = content.split('\n')
            in_table = False
            
            for line in lines:
                line = line.strip()
                
                # Look for table start
                if line.startswith('| Attribute | Value |'):
                    in_table = True
                    continue
                elif line.startswith('|---') and in_table:
                    continue
                elif in_table and line.startswith('|') and '|' in line[1:]:
                    # Parse table row
                    parts = [part.strip().strip('*') for part in line.split('|')[1:-1]]
                    if len(parts) >= 2:
                        attribute = parts[0].lower()
                        value = parts[1].replace('\\|', '|')  # Unescape pipes
                        
                        if 'requirement id' in attribute:
                            req_data['requirement_id'] = value
                        elif 'category' in attribute or 'functional activity' in attribute:
                            req_data['category'] = value
                        elif 'topic' in attribute:
                            req_data['topic'] = value
                        elif 'short description' in attribute:
                            req_data['short_description'] = value
                        elif 'description overview' in attribute:
                            req_data['description'] = value
                        elif 'priority' in attribute:
                            req_data['priority'] = value
                elif in_table and not line.startswith('|'):
                    # End of table
                    break
            
            return req_data
            
        except Exception as e:
            self.log(f"Error reading file {filepath}: {e}")
            return None
    
    def get_all_requirement_files(self):
        """Get all .md files in the vault directory"""
        output_dir = self.output_dir.get()
        
        if not os.path.exists(output_dir):
            return []
            
        md_files = []
        for filename in os.listdir(output_dir):
            if filename.endswith('.md') and not filename.startswith('0_'):  # Skip overview files
                filepath = os.path.join(output_dir, filename)
                req_data = self.extract_requirement_data_from_file(filepath)
                if req_data:
                    md_files.append(req_data)
        
        return md_files
    
    def generate_overview_only(self):
        """Generate overview file of all requirements"""
        if not self.output_dir.get():
            messagebox.showerror("Error", "Please select Obsidian vault directory first")
            return
        
        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("Error", "Obsidian vault directory does not exist")
            return
        
        try:
            self.status_text.set("Generating overview...")
            self.log("=" * 50)
            self.log("GENERATING REQUIREMENTS OVERVIEW")
            
            # Get all requirement files
            all_requirements = self.get_all_requirement_files()
            
            if not all_requirements:
                self.log("⚠ No requirement files found in vault")
                self.status_text.set("No requirements found")
                messagebox.showwarning("No Requirements", 
                    "No requirement files found in the selected vault.\n"
                    "Use 'Create Missing Files' first to create requirement files.")
                return
            
            # Sort requirements by ID, then by topic
            def sort_key(req):
                return (req['requirement_id'] or 'zzz', req['topic'] or 'zzz')
            
            all_requirements.sort(key=sort_key)
            
            # Create overview content
            content = []
            content.append("# Requirements Overview")
            content.append("")
            content.append(f"*Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}*")
            content.append("")
            
            # Summary statistics
            content.append("## Summary")
            content.append("")
            content.append(f"- **Total Requirements**: {len(all_requirements)}")
            content.append("")
            
            # Category breakdown
            categories = {}
            priority_levels = {}
            
            for req in all_requirements:
                cat = req['category'] or 'Uncategorized'
                pri = req['priority'] or 'Not specified'
                categories[cat] = categories.get(cat, 0) + 1
                priority_levels[pri] = priority_levels.get(pri, 0) + 1
            
            if categories:
                content.append("### Categories")
                for cat, count in sorted(categories.items()):
                    content.append(f"- **{cat}**: {count} requirements")
                content.append("")
            
            if priority_levels:
                content.append("### Priority Levels")
                for pri, count in sorted(priority_levels.items()):
                    content.append(f"- **{pri}**: {count} requirements")
                content.append("")
            
            # Requirements table
            content.append("## All Requirements")
            content.append("")
            content.append("| ID | Category | Topic | Short Description | Description Overview | Priority | File |")
            content.append("|:---|:---------|:------|:------------------|:---------------------|:---------|:-----|")
            
            for req in all_requirements:
                # Truncate long descriptions for table readability
                desc = req['description'][:100] + "..." if len(req['description']) > 100 else req['description']
                # Escape pipes in cell content
                desc = desc.replace('|', '\\|')
                category = req['category'].replace('|', '\\|')
                topic = req['topic'].replace('|', '\\|')
                short_description = req['short_description'].replace('|', '\\|')
                priority = req['priority'].replace('|', '\\|')
                
                # Create link to file (remove .md extension for cleaner links)
                file_link = f"[[{req['filename'][:-3]}]]" if req['filename'].endswith('.md') else f"[[{req['filename']}]]"
                
                content.append(f"| {req['requirement_id']} | {category} | {topic} | {short_description} | {desc} | {priority} | {file_link} |")
            
            content.append("")
            
            # Usage instructions
            content.append("---")
            content.append("")
            content.append("## Usage Instructions")
            content.append("")
            content.append("This overview provides a comprehensive view of all requirements. To analyze specific requirements:")
            content.append("")
            content.append("1. **Browse by Category**: Look for patterns in similar functional areas")
            content.append("2. **Filter by Priority**: Focus on high-priority requirements first")
            content.append("3. **Follow Links**: Click on any file link to view detailed requirement information")
            content.append("4. **Search & Filter**: Use Obsidian's search to find requirements by keywords")
            content.append("")
            
            # Write overview file
            overview_path = os.path.join(self.output_dir.get(), "0_Requirements_Overview.md")
            with open(overview_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(content))
            
            self.log(f"✓ Created overview file: 0_Requirements_Overview.md")
            self.log(f"  - {len(all_requirements)} requirements included")
            
            self.status_text.set("Overview generated!")
            
            # Show success message
            success_msg = (f"Overview Generated Successfully!\n\n"
                          f"Total Requirements: {len(all_requirements)}\n"
                          f"Categories: {len(categories)}\n\n"
                          f"Overview file: 0_Requirements_Overview.md")
            
            messagebox.showinfo("Overview Complete", success_msg)
            self.log("=" * 50)
            
        except Exception as e:
            error_msg = f"Error generating overview: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.status_text.set("Error occurred")
            messagebox.showerror("Error", error_msg)
            
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = RequirementsConverter()
    app.run()