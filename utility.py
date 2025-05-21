import os
import re
import shutil
import csv
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter

class MultiUtilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Utility Application")
        self.root.geometry("900x700")

        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create tabs
        self.create_folder_utilities_tab()
        self.create_csv_utilities_tab()
        self.create_pdf_utilities_tab()
        self.create_file_search_copy_tab()
        self.create_csv_viewer_tab()

    # ---------------- Folder Utilities Tab ----------------
    def create_folder_utilities_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Folder Utilities")

        btn_remove_subfolders = ttk.Button(tab, text="Remove Empty Subfolders", command=self.remove_empty_subfolders)
        btn_remove_subfolders.pack(pady=10, fill=tk.X)

        btn_remove_folders = ttk.Button(tab, text="Remove Empty Folders", command=self.remove_empty_folders)
        btn_remove_folders.pack(pady=10, fill=tk.X)

        btn_browse_folder = ttk.Button(tab, text="Browse Folder Contents", command=self.browse_folder_contents)
        btn_browse_folder.pack(pady=10, fill=tk.X)

        self.folder_contents_tree = ttk.Treeview(tab, columns=("Name", "Type"), show="headings")
        self.folder_contents_tree.heading("Name", text="Name")
        self.folder_contents_tree.heading("Type", text="Type")
        self.folder_contents_tree.column("Name", width=600)
        self.folder_contents_tree.column("Type", width=100)
        self.folder_contents_tree.pack(fill=tk.BOTH, expand=True, pady=10)

    def remove_empty_subfolders(self):
        folder = filedialog.askdirectory(title="Select Parent Folder")
        if not folder:
            return
        def task():
            count = 0
            for root, dirs, _ in os.walk(folder, topdown=False):
                for d in dirs:
                    path = os.path.join(root, d)
                    if not os.listdir(path):
                        try:
                            os.rmdir(path)
                            count += 1
                        except Exception as e:
                            print(f"Failed to remove {path}: {e}")
            messagebox.showinfo("Done", f"Removed {count} empty subfolders (excluding root).")
        threading.Thread(target=task, daemon=True).start()

    def remove_empty_folders(self):
        folder = filedialog.askdirectory(title="Select Folder to Remove Empty Folders")
        if not folder:
            return
        def recursive_remove(path):
            count = 0
            if not os.path.isdir(path):
                return count
            for entry in os.listdir(path):
                full_path = os.path.join(path, entry)
                if os.path.isdir(full_path):
                    count += recursive_remove(full_path)
            if not os.listdir(path):
                try:
                    os.rmdir(path)
                    count += 1
                except Exception as e:
                    print(f"Failed to remove {path}: {e}")
            return count
        def task():
            count = recursive_remove(folder)
            messagebox.showinfo("Done", f"Removed {count} empty folders (including root if empty).")
        threading.Thread(target=task, daemon=True).start()

    def browse_folder_contents(self):
        folder = filedialog.askdirectory(title="Select Folder to Browse")
        if not folder:
            return
        try:
            items = os.listdir(folder)
            self.folder_contents_tree.delete(*self.folder_contents_tree.get_children())
            for item in items:
                path = os.path.join(folder, item)
                typ = "Folder" if os.path.isdir(path) else "File"
                self.folder_contents_tree.insert("", "end", values=(item, typ))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list folder contents:\n{e}")

    # ---------------- CSV Utilities Tab ----------------
    def create_csv_utilities_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="CSV Utilities")

        btn_save_paths = ttk.Button(tab, text="Save Subfolder Paths to CSV", command=self.save_subfolder_paths_to_csv)
        btn_save_paths.pack(pady=10, fill=tk.X)

        btn_save_names = ttk.Button(tab, text="Save Subfolder Names to CSV", command=self.save_subfolder_names_to_csv)
        btn_save_names.pack(pady=10, fill=tk.X)

        btn_create_folders = ttk.Button(tab, text="Create Folders From CSV", command=self.create_folders_from_csv)
        btn_create_folders.pack(pady=10, fill=tk.X)

    def save_subfolder_paths_to_csv(self):
        folder = filedialog.askdirectory(title="Select Parent Folder")
        if not folder:
            return
        csv_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                                title="Save Subfolder Paths As")
        if not csv_path:
            return
        def task():
            paths = []
            for root, dirs, _ in os.walk(folder):
                for d in dirs:
                    paths.append([os.path.join(root, d)])
            try:
                with open(csv_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Subfolder Path"])
                    writer.writerows(paths)
                messagebox.showinfo("Success", f"Saved {len(paths)} subfolder paths to CSV.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV:\n{e}")
        threading.Thread(target=task, daemon=True).start()

    def save_subfolder_names_to_csv(self):
        folder = filedialog.askdirectory(title="Select Parent Folder")
        if not folder:
            return
        csv_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                                title="Save Subfolder Names As")
        if not csv_path:
            return
        def task():
            names = []
            for root, dirs, _ in os.walk(folder):
                for d in dirs:
                    names.append([d])
            try:
                with open(csv_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Subfolder Name"])
                    writer.writerows(names)
                messagebox.showinfo("Success", f"Saved {len(names)} subfolder names to CSV.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV:\n{e}")
        threading.Thread(target=task, daemon=True).start()

    def create_folders_from_csv(self):
        csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], title="Select CSV with Folder Names")
        if not csv_path:
            return
        root_folder = filedialog.askdirectory(title="Select Root Folder to Create Folders In")
        if not root_folder:
            return
        def task():
            created = 0
            try:
                with open(csv_path, "r", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    header = next(reader, None)
                    for row in reader:
                        if row:
                            folder_name = row[0].strip()
                            if folder_name:
                                folder_path = os.path.join(root_folder, folder_name)
                                if not os.path.exists(folder_path):
                                    os.makedirs(folder_path)
                                    created += 1
                messagebox.showinfo("Success", f"Created {created} folders.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create folders:\n{e}")
        threading.Thread(target=task, daemon=True).start()

    # ---------------- PDF Utilities Tab ----------------
    def create_pdf_utilities_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="PDF Utilities")

        btn_split_pdf = ttk.Button(tab, text="Split PDF by Pages", command=self.split_pdf_dialog)
        btn_split_pdf.pack(pady=10, fill=tk.X)

        btn_merge_pdf = ttk.Button(tab, text="Merge PDFs in Pairs", command=self.merge_pdfs_dialog)
        btn_merge_pdf.pack(pady=10, fill=tk.X)

    def split_pdf_dialog(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")], title="Select PDF to Split")
        if not pdf_path:
            return
        output_dir = filedialog.askdirectory(title="Select Output Folder for Split PDFs")
        if not output_dir:
            return
        def on_submit():
            try:
                pages_per_split = int(pages_entry.get())
                if pages_per_split <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Pages per split must be a positive integer.")
                return
            split_window.destroy()
            threading.Thread(target=self.split_pdf, args=(pdf_path, output_dir, pages_per_split), daemon=True).start()
        split_window = tk.Toplevel(self.root)
        split_window.title("Pages per Split")
        ttk.Label(split_window, text="Pages per split:").pack(padx=10, pady=5)
        pages_entry = ttk.Entry(split_window)
        pages_entry.pack(padx=10, pady=5)
        pages_entry.insert(0, "1")
        ttk.Button(split_window, text="Start Splitting", command=on_submit).pack(pady=10)

    def sanitize_filename(self, text):
        text = text.strip()
        text = re.sub(r'[^\w\- ]', '', text)
        text = re.sub(r'\s+', '_', text)
        return text or "page"

    def extract_header_text(self, pdf_reader, page_num=0, max_chars=30):
        try:
            page = pdf_reader.pages[page_num]
            text = page.extract_text() or ""
            first_line = text.strip().split('\n')[0]
            return first_line[:max_chars]
        except Exception:
            return "page"

    def split_pdf(self, pdf_path, output_dir, pages_per_split):
        try:
            pdf_reader = PdfReader(pdf_path)
            total_pages = len(pdf_reader.pages)
            count = 0
            for start in range(0, total_pages, pages_per_split):
                writer = PdfWriter()
                end = min(start + pages_per_split, total_pages)
                for i in range(start, end):
                    writer.add_page(pdf_reader.pages[i])
                header_text = self.extract_header_text(pdf_reader, page_num=start)
                filename = f"{self.sanitize_filename(header_text)}_{start + 1}_to_{end}.pdf"
                output_path = os.path.join(output_dir, filename)
                with open(output_path, 'wb') as f:
                    writer.write(f)
                count += 1
                self.update_status(f"Split pages {start + 1} to {end} into {filename}")
            messagebox.showinfo("Done", f"Splitting complete! Created {count} files.")
        except Exception as e:
            messagebox.showerror("Error", f"Error during splitting:\n{e}")

    def merge_pdfs_dialog(self):
        folder = filedialog.askdirectory(title="Select Folder Containing PDFs to Merge")
        if not folder:
            return
        output_folder = filedialog.askdirectory(title="Select Output Folder for Merged PDFs")
        if not output_folder:
            return
        threading.Thread(target=self.merge_pdfs_in_pairs, args=(folder, output_folder), daemon=True).start()

    def merge_pdfs_in_pairs(self, folder, output_folder):
        try:
            pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
            pdf_files.sort()
            if not pdf_files:
                messagebox.showinfo("Info", "No PDF files found in input folder.")
                return
            os.makedirs(output_folder, exist_ok=True)
            for i in range(0, len(pdf_files), 2):
                writer = PdfWriter()
                pair = pdf_files[i:i+2]
                first_pdf_path = os.path.join(folder, pair[0])
                with open(first_pdf_path, 'rb') as f:
                    reader = PdfReader(f)
                    header_text = self.extract_header_text(reader)
                for pdf_file in pair:
                    pdf_path = os.path.join(folder, pdf_file)
                    with open(pdf_path, 'rb') as f:
                        reader = PdfReader(f)
                        for page in reader.pages:
                            writer.add_page(page)
                output_filename = f"{self.sanitize_filename(header_text)}_merged_{i//2 + 1}.pdf"
                output_path = os.path.join(output_folder, output_filename)
                with open(output_path, 'wb') as out_pdf:
                    writer.write(out_pdf)
                self.update_status(f"Merged pair {i+1} and {i+2 if i+1 < len(pdf_files) else ''} into {output_filename}")
            messagebox.showinfo("Done", "Merging complete!")
        except Exception as e:
            messagebox.showerror("Error", f"Error during merging:\n{e}")

    # ---------------- File Search & Copy Tab ----------------
    def create_file_search_copy_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="File Search & Copy")

        btn_search_files = ttk.Button(tab, text="Search Files by Extension & Save CSV", command=self.search_files_by_ext_dialog)
        btn_search_files.pack(pady=10, fill=tk.X)

        btn_filter_copy = ttk.Button(tab, text="Filter CSV & Copy DJI Images", command=self.filter_csv_and_copy_images_dialog)
        btn_filter_copy.pack(pady=10, fill=tk.X)

    def search_files_by_ext_dialog(self):
        folder = filedialog.askdirectory(title="Select Folder to Search")
        if not folder:
            return
        def on_submit():
            ext = ext_entry.get().strip()
            if not ext.startswith('.'):
                ext = '.' + ext
            csv_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                                    title="Save CSV File")
            if not csv_path:
                return
            search_window.destroy()
            threading.Thread(target=self.search_files_by_extension, args=(folder, ext, csv_path), daemon=True).start()

        search_window = tk.Toplevel(self.root)
        search_window.title("File Extension to Search")
        ttk.Label(search_window, text="File Extension (e.g. .pdf):").pack(padx=10, pady=5)
        ext_entry = ttk.Entry(search_window)
        ext_entry.pack(padx=10, pady=5)
        ext_entry.insert(0, ".pdf")
        ttk.Button(search_window, text="Start Search", command=on_submit).pack(pady=10)

    def search_files_by_extension(self, folder, ext, csv_path):
        found_files = []
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.lower().endswith(ext.lower()):
                    found_files.append(os.path.join(root, file))
        try:
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["File Path"])
                for filepath in found_files:
                    writer.writerow([filepath])
            messagebox.showinfo("Success", f"Found {len(found_files)} files.\nCSV saved at:\n{csv_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write CSV:\n{e}")

    def filter_csv_and_copy_images_dialog(self):
        csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], title="Select CSV File to Filter")
        if not csv_file_path:
            return
        threading.Thread(target=self.filter_csv_and_copy_images, args=(csv_file_path,), daemon=True).start()

    def filter_csv_and_copy_images(self, csv_file_path):
        missing_images = []
        output_folder = os.path.join(os.path.dirname(csv_file_path), 'output')
        os.makedirs(output_folder, exist_ok=True)
        new_csv_path = os.path.join(os.path.dirname(csv_file_path), 'filtered_output.csv')
        try:
            with open(csv_file_path, newline='', encoding='utf-8') as csvfile, \
                 open(new_csv_path, mode='w', newline='', encoding='utf-8') as new_csv:
                csv_reader = csv.reader(csvfile)
                csv_writer = csv.writer(new_csv)
                rows = list(csv_reader)
                if len(rows) < 2:
                    messagebox.showinfo("Info", "CSV file has less than 2 rows.")
                    return
                csv_writer.writerow(['Column1', 'Column2', 'Filename'])
                for row in rows:
                    if len(row) < 3:
                        continue
                    col1, col2, filename = row[0], row[1], row[2]
                    try:
                        col1_value = float(col1)
                        col2_value = float(col2)
                    except ValueError:
                        continue
                    if 60 <= col1_value <= 70 and 60 <= col2_value <= 70:
                        if filename.startswith('DJI') and filename.endswith('jpeg'):
                            csv_writer.writerow([col1, col2, filename])
                            image_path = os.path.join(os.path.dirname(csv_file_path), filename)
                            if os.path.exists(image_path):
                                shutil.copy(image_path, output_folder)
                            else:
                                missing_images.append(image_path)
                if missing_images:
                    for image_path in missing_images:
                        if os.path.exists(image_path):
                            shutil.copy(image_path, output_folder)
                messagebox.showinfo("Success", f"Filtering and copying complete.\nFiltered CSV saved at:\n{new_csv_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error processing CSV file:\n{e}")

    # ---------------- CSV Viewer Tab ----------------
    def create_csv_viewer_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="CSV Viewer")

        btn_load_csv = ttk.Button(tab, text="Load CSV File", command=self.load_csv_file)
        btn_load_csv.pack(pady=10, fill=tk.X)

        self.csv_tree = ttk.Treeview(tab)
        self.csv_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(self.csv_tree, orient=tk.VERTICAL, command=self.csv_tree.yview)
        self.csv_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def load_csv_file(self):
        csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], title="Select CSV File")
        if not csv_file_path:
            return
        try:
            with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader, None)
                self.csv_tree.delete(*self.csv_tree.get_children())
                if header:
                    self.csv_tree["columns"] = header
                    for col in header:
                        self.csv_tree.heading(col, text=col)
                        self.csv_tree.column(col, width=150)
                else:
                    self.csv_tree["columns"] = ("Data",)
                    self.csv_tree.heading("Data", text="Data")
                    self.csv_tree.column("Data", width=300)
                for row in reader:
                    self.csv_tree.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MultiUtilityApp(root)
    root.mainloop()
