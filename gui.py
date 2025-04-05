import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import openpyxl
import os
import openpyxl
from config import ClassConfig, CategoryConfig
from excel_generator import generate_grade_file

class CategoryFrame(ttk.Frame):
    def __init__(self, parent, category=None, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.category = category
        
        # Category Name
        ttk.Label(self, text="Category Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.name_var = tk.StringVar(value=category.name if category else "")
        ttk.Entry(self, textvariable=self.name_var, width=30).grid(row=0, column=1, padx=5, pady=2)
        
        # Weight
        ttk.Label(self, text="Weight (%):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.weight_var = tk.DoubleVar(value=category.weight if category else 0)
        ttk.Spinbox(self, from_=0, to=100, textvariable=self.weight_var, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        # Total Items
        ttk.Label(self, text="Total Items:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.total_var = tk.IntVar(value=category.total_items if category else 1)
        ttk.Spinbox(self, from_=1, to=100, textvariable=self.total_var, width=10).grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        # Best Of
        ttk.Label(self, text="Best Of:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.best_var = tk.IntVar(value=category.best_of if category else 1)
        ttk.Spinbox(self, from_=1, to=100, textvariable=self.best_var, width=10).grid(row=3, column=1, sticky="w", padx=5, pady=2)
        
        # Delete button
        ttk.Button(self, text="Remove", command=self.remove).grid(row=4, column=0, columnspan=2, pady=5)
        
        # Configure frame
        self.config(padding="10", relief="groove", borderwidth=1)
    
    def remove(self):
        self.destroy()
    
    def get_category(self):
        try:
            return CategoryConfig(
                name=self.name_var.get(),
                weight=float(self.weight_var.get()),
                total_items=int(self.total_var.get()),
                best_of=int(self.best_var.get())
            )
        except ValueError:
            return None

class GradeCalculatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Grade Calculator Generator")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        # Main container
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Class Info Frame
        class_frame = ttk.LabelFrame(main_frame, text="Class Information", padding="10")
        class_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(class_frame, text="Class Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.class_name_var = tk.StringVar()
        ttk.Entry(class_frame, textvariable=self.class_name_var, width=30).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        # Categories Frame
        categories_container = ttk.LabelFrame(main_frame, text="Assessment Categories", padding="10")
        categories_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollable frame for categories
        self.canvas = tk.Canvas(categories_container)
        scrollbar = ttk.Scrollbar(categories_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add Category Button
        ttk.Button(main_frame, text="Add Category", command=self.add_category).pack(pady=5)
        
        # Action Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Generate Excel", command=self.generate_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Load MAT137 Template", command=self.load_mat137).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Load Existing Excel", command=self.load_existing_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.quit).pack(side=tk.RIGHT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)
        
        self.category_frames = []
        self.existing_wb_path = None
    
    def add_category(self, category=None):
        frame = CategoryFrame(self.scrollable_frame, category=category)
        frame.pack(fill=tk.X, pady=5, padx=5)
        self.category_frames.append(frame)
        return frame
    
    def load_mat137(self):
        # Clear existing categories
        for frame in self.category_frames:
            frame.destroy()
        self.category_frames.clear()
        
        # Set class name
        self.class_name_var.set("MAT137")
        
        # Add MAT137 categories
        categories = [
            CategoryConfig("Pre-course Module", 2, 1, 1),
            CategoryConfig("Computation Quizzes", 3, 1, 1),
            CategoryConfig("Tutorial Worksheets", 5, 23, 17),
            CategoryConfig("Pre-class Quizzes", 5, 70, 50),
            CategoryConfig("Problem Sets", 12, 8, 6),
            CategoryConfig("Term Tests", 39, 4, 3),
            CategoryConfig("Final Exam", 34, 1, 1)
        ]
        
        for category in categories:
            self.add_category(category)
        
        self.status_var.set("MAT137 template loaded")
    
    def validate_config(self):
        # Check class name
        class_name = self.class_name_var.get().strip()
        if not class_name:
            messagebox.showerror("Error", "Please enter a class name")
            return None
        
        # Get categories
        categories = []
        total_weight = 0
        
        for frame in self.category_frames:
            if not frame.winfo_exists():  # Skip destroyed frames
                continue
                
            category = frame.get_category()
            if not category:
                messagebox.showerror("Error", "Invalid category data")
                return None
                
            if not category.name.strip():
                messagebox.showerror("Error", "Category name cannot be empty")
                return None
                
            if category.best_of > category.total_items:
                messagebox.showerror("Error", f"'Best of' ({category.best_of}) cannot be greater than 'Total items' ({category.total_items}) for {category.name}")
                return None
            
            total_weight += category.weight
            categories.append(category)
        
        if not categories:
            messagebox.showerror("Error", "Please add at least one category")
            return None
            
        # Check total weight
        if abs(total_weight - 100) > 0.01:  # Allow small floating point errors
            result = messagebox.askquestion("Warning", 
                f"Total weight is {total_weight}%, not 100%. Continue anyway?")
            if result != "yes":
                return None
        
        return ClassConfig(class_name=class_name, categories=categories)
    
    def load_existing_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Select existing grade calculator file"
        )
        if file_path:
            self.existing_wb_path = file_path
            self.status_var.set(f"Loaded existing file: {file_path}")
    
    def generate_excel(self):
        config = self.validate_config()
        if not config:
            return
            
        try:
            if self.existing_wb_path:
                # Load existing workbook
                wb = openpyxl.load_workbook(self.existing_wb_path)
                wb = generate_grade_file(config, wb)
                default_name = os.path.basename(self.existing_wb_path)
            else:
                # Create new workbook
                wb = generate_grade_file(config)
                default_name = f"{config.class_name}_Grade_Calculator.xlsx"

            # Get save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_name
            )
            
            if not file_path:  # User cancelled
                return
                
            # Save the workbook
            wb.save(file_path)
            
            # Clear existing workbook reference after save
            self.existing_wb_path = None
                
            self.status_var.set(f"Excel file saved: {file_path}")
            messagebox.showinfo("Success", f"Grade calculator saved successfully as:\n{file_path}")
            os.system(f"open {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create Excel file: {str(e)}")
            self.status_var.set("Error creating Excel file")

if __name__ == "__main__":
    app = GradeCalculatorApp()
    app.mainloop()
