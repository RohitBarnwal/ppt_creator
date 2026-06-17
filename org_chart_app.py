import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Import our backend generation module
try:
    import build_org_structure
except ImportError:
    # Fallback in case of path issues
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import build_org_structure

# Set the default theme and color options
ctk.set_appearance_mode("System")  # System theme (dark/light)
ctk.set_default_color_theme("blue") # Theme color

class OrgChartApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure Main Window
        self.title("Paytm Organization Chart Generator")
        self.geometry("520x420")
        self.resizable(False, False)
        
        # Internal State
        self.selected_excel_path = ""
        
        # Layout Definition
        self.create_widgets()
        
    def create_widgets(self):
        # 1. Main Title Header Block
        self.header_label = ctk.CTkLabel(
            self, 
            text="Org Chart Creator Utility", 
            font=ctk.CTkFont(family="Arial", size=22, weight="bold")
        )
        self.header_label.pack(pady=(25, 5))
        
        self.desc_label = ctk.CTkLabel(
            self, 
            text="Generate connected multi-level PowerPoint org charts dynamically from an Excel file.",
            font=ctk.CTkFont(family="Arial", size=12),
            text_color="gray"
        )
        self.desc_label.pack(pady=(0, 20))
        
        # 2. File Selection Panel (Frame)
        self.panel_frame = ctk.CTkFrame(self, width=460, height=220)
        self.panel_frame.pack(padx=30, pady=10, fill="both", expand=True)
        
        # Upload Button
        self.upload_btn = ctk.CTkButton(
            self.panel_frame, 
            text="Browse Excel File", 
            command=self.browse_excel,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.upload_btn.pack(pady=(20, 10))
        
        # Label to show the selected file path
        self.file_label = ctk.CTkLabel(
            self.panel_frame, 
            text="No file selected", 
            font=ctk.CTkFont(size=11, slant="italic"),
            text_color="gray",
            wraplength=400
        )
        self.file_label.pack(pady=(0, 15))
        
        # 3. Output Custom File Name Section
        self.output_label = ctk.CTkLabel(
            self.panel_frame, 
            text="Output PowerPoint Name:", 
            font=ctk.CTkFont(size=12)
        )
        self.output_label.pack(anchor="w", padx=30, pady=(5, 2))
        
        self.output_entry = ctk.CTkEntry(
            self.panel_frame, 
            width=380, 
            placeholder_text="Enter output name (e.g., final_org_chart.pptx)"
        )
        self.output_entry.pack(padx=30, pady=(0, 20))
        
        # 4. Status Indicator (Progress Bar & Label)
        self.status_label = ctk.CTkLabel(
            self, 
            text="Status: Ready", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1f77b4"
        )
        self.status_label.pack(pady=(10, 5))
        
        self.progress_bar = ctk.CTkProgressBar(self, width=460)
        self.progress_bar.pack(pady=(0, 15))
        self.progress_bar.set(0.0)
        
        # 5. Bottom Generation Actions
        self.generate_btn = ctk.CTkButton(
            self, 
            text="Generate Organization Chart", 
            command=self.start_generation_thread,
            state="disabled", # Disabled until file is loaded
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.generate_btn.pack(pady=(0, 25))

    def browse_excel(self):
        # Open browser dialog
        file_path = filedialog.askopenfilename(
            title="Select Input Excel file",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.selected_excel_path = file_path
            self.file_label.configure(text=os.path.basename(file_path), text_color=["black", "white"])
            self.generate_btn.configure(state="normal")
            
            # Auto-populate output name if empty
            default_out = os.path.join(
                os.path.dirname(file_path), 
                "Automated_Org_Structure_" + os.path.basename(file_path).replace(".xlsx", ".pptx").replace(".xls", ".pptx")
            )
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, default_out)
            
            self.update_status("File loaded. Ready to generate.", "#2ca02c")
            
    def update_status(self, text, color):
        self.status_label.configure(text=f"Status: {text}", text_color=color)
        
    def start_generation_thread(self):
        # Prevent double clicks
        self.generate_btn.configure(state="disabled")
        self.upload_btn.configure(state="disabled")
        self.output_entry.configure(state="disabled")
        
        # Set progress bar to pulse / active mode
        self.progress_bar.set(0.2)
        self.update_status("Processing Dataset...", "#e377c2")
        
        # Fire background thread
        thread = threading.Thread(target=self.run_generation)
        thread.daemon = True
        thread.start()
        
    def run_generation(self):
        try:
            excel_path = self.selected_excel_path
            out_pptx = self.output_entry.get().strip()
            
            if not out_pptx:
                # Fallback to default
                out_pptx = excel_path.replace(".xlsx", ".pptx").replace(".xls", ".pptx")
                
            if not out_pptx.endswith(".pptx"):
                out_pptx += ".pptx"
                
            self.progress_bar.set(0.6)
            
            # Execute backend generator
            success = build_org_structure.build_org_presentation(excel_path, out_pptx)
            
            if success:
                self.progress_bar.set(1.0)
                self.update_status("Generation Completed Successfully!", "#2ca02c")
                messagebox.showinfo("Success", f"Org Chart PPTX has been generated successfully!\n\nSaved to:\n{out_pptx}")
            else:
                self.progress_bar.set(0.0)
                self.update_status("Generation Failed", "#d62728")
                messagebox.showerror("Error", "Generation failed. Please verify that the Excel sheet is correctly formatted.")
        except Exception as e:
            self.progress_bar.set(0.0)
            self.update_status("System Error", "#d62728")
            messagebox.showerror("System Error", f"An unexpected error occurred:\n{str(e)}")
        finally:
            # Re-enable interactive widgets
            self.generate_btn.configure(state="normal")
            self.upload_btn.configure(state="normal")
            self.output_entry.configure(state="normal")

if __name__ == "__main__":
    app = OrgChartApp()
    app.mainloop()
