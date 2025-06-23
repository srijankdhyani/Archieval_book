import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import cv2
import pytesseract
from PIL import Image, ImageTk
import os
import sys
from datetime import datetime
import threading
import numpy as np

# PDF and document handling
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# Scanner support (optional - may need specific installation)
try:
    import twain
    TWAIN_AVAILABLE = True
except ImportError:
    TWAIN_AVAILABLE = False
    print("Warning: pytwain not available. Scanner functionality will be limited.")

class OCRApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced OCR Application")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Configure tesseract path
        pytesseract.pytesseract.tesseract_cmd = r"E:\Github\Archieval_Sofware\Tesseract\tesseract.exe"
        
        # Variables
        self.current_image = None
        self.extracted_text = ""
        self.camera = None
        self.camera_active = False
        
        # Setup GUI
        self.setup_gui()
        
    def setup_gui(self):
        """Setup the main GUI interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Advanced OCR Application", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input methods frame
        input_frame = ttk.LabelFrame(main_frame, text="Input Methods", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Input buttons
        ttk.Button(input_frame, text="📷 Camera Capture", 
                  command=self.open_camera, width=20).grid(row=0, column=0, padx=5)
        
        ttk.Button(input_frame, text="🖼️ Select Files", 
                  command=self.select_files, width=20).grid(row=0, column=1, padx=5)
        
        ttk.Button(input_frame, text="📁 Batch Process", 
                  command=self.batch_process, width=20).grid(row=0, column=2, padx=5)
        
        if TWAIN_AVAILABLE:
            ttk.Button(input_frame, text="🖨️ Scanner", 
                      command=self.scan_document, width=20).grid(row=0, column=3, padx=5)
        
        # Image display frame
        image_frame = ttk.LabelFrame(main_frame, text="Image Preview", padding="10")
        image_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        image_frame.columnconfigure(0, weight=1)
        image_frame.rowconfigure(0, weight=1)
        
        # Image label with scrollbars
        self.image_canvas = tk.Canvas(image_frame, bg='white', width=400, height=300)
        self.image_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # OCR results frame
        text_frame = ttk.LabelFrame(main_frame, text="OCR Results", padding="10")
        text_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # Text display
        self.text_display = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, 
                                                     width=50, height=20)
        self.text_display.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Control buttons frame
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(control_frame, text="🔍 Run OCR", 
                  command=self.run_ocr, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="💾 Save as Text", 
                  command=self.save_as_text).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="📄 Save as PDF", 
                  command=self.save_as_pdf).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="🗑️ Clear", 
                  command=self.clear_all).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select an input method to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_var.set(message)
        self.root.update()
    
    def display_image(self, image_path_or_array):
        """Display image in the canvas"""
        try:
            if isinstance(image_path_or_array, str):
                # Load from file path
                image = Image.open(image_path_or_array)
            else:
                # Convert from numpy array (OpenCV format)
                if len(image_path_or_array.shape) == 3:
                    # Convert BGR to RGB
                    image_rgb = cv2.cvtColor(image_path_or_array, cv2.COLOR_BGR2RGB)
                    image = Image.fromarray(image_rgb)
                else:
                    image = Image.fromarray(image_path_or_array)
            
            # Resize image to fit canvas
            canvas_width = self.image_canvas.winfo_width()
            canvas_height = self.image_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                image.thumbnail((canvas_width-20, canvas_height-20), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage and display
            photo = ImageTk.PhotoImage(image)
            self.image_canvas.delete("all")
            self.image_canvas.create_image(canvas_width//2, canvas_height//2, 
                                         image=photo, anchor=tk.CENTER)
            self.image_canvas.image = photo  # Keep a reference
            
            self.current_image = image
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display image: {str(e)}")
    
    def open_camera(self):
        """Open camera for capturing images"""
        def camera_thread():
            self.camera = cv2.VideoCapture(0)
            if not self.camera.isOpened():
                messagebox.showerror("Error", "Could not open camera")
                return
            
            self.camera_active = True
            camera_window = tk.Toplevel(self.root)
            camera_window.title("Camera Capture")
            camera_window.geometry("640x480")
            
            camera_label = tk.Label(camera_window)
            camera_label.pack(expand=True)
            
            button_frame = tk.Frame(camera_window)
            button_frame.pack(side=tk.BOTTOM, pady=10)
            
            def capture_image():
                ret, frame = self.camera.read()
                if ret:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"camera_capture_{timestamp}.jpg"
                    cv2.imwrite(filename, frame)
                    self.display_image(frame)
                    self.update_status(f"Image captured: {filename}")
                    camera_window.destroy()
                    self.camera_active = False
                    self.camera.release()
            
            def close_camera():
                self.camera_active = False
                camera_window.destroy()
                if self.camera:
                    self.camera.release()
            
            tk.Button(button_frame, text="📸 Capture", command=capture_image,
                     bg='green', fg='white', font=('Arial', 12)).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="❌ Close", command=close_camera,
                     bg='red', fg='white', font=('Arial', 12)).pack(side=tk.LEFT, padx=10)
            
            camera_window.protocol("WM_DELETE_WINDOW", close_camera)
            
            def update_camera():
                if self.camera_active and self.camera.isOpened():
                    ret, frame = self.camera.read()
                    if ret:
                        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                        image = Image.fromarray(frame_rgb)
                        image = image.resize((640, 480), Image.Resampling.LANCZOS)
                        photo = ImageTk.PhotoImage(image)
                        camera_label.configure(image=photo)
                        camera_label.image = photo
                    camera_window.after(30, update_camera)
            
            update_camera()
        
        threading.Thread(target=camera_thread, daemon=True).start()
    
    def select_files(self):
        """Select individual files for OCR"""
        file_paths = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.tiff *.bmp *.gif"),
                ("All files", "*.*")
            ]
        )
        
        if file_paths:
            if len(file_paths) == 1:
                self.display_image(file_paths[0])
                self.update_status(f"Selected: {os.path.basename(file_paths[0])}")
            else:
                self.process_multiple_files(file_paths)
    
    def batch_process(self):
        """Process all images in a selected folder"""
        folder_path = filedialog.askdirectory(title="Select Folder for Batch Processing")
        
        if folder_path:
            image_extensions = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')
            image_files = []
            
            for filename in os.listdir(folder_path):
                if filename.lower().endswith(image_extensions):
                    image_files.append(os.path.join(folder_path, filename))
            
            if image_files:
                self.process_multiple_files(image_files)
            else:
                messagebox.showinfo("Info", "No image files found in the selected folder.")
    
    def scan_document(self):
        """Scan document using TWAIN interface"""
        if not TWAIN_AVAILABLE:
            messagebox.showerror("Error", "TWAIN scanner interface not available.")
            return
        
        try:
            # Initialize TWAIN
            sm = twain.SourceManager(0)
            sources = sm.GetSourceList()
            
            if not sources:
                messagebox.showerror("Error", "No scanners found.")
                return
            
            # Use first available scanner
            scanner = sm.OpenSource(sources[0])
            scanner.RequestAcquire(0, 0)  # Show scanner interface
            
            # This is a simplified implementation
            # In practice, you'd need to handle TWAIN callbacks properly
            messagebox.showinfo("Scanner", "Please use your scanner software to scan the document, then select the scanned file.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Scanner error: {str(e)}")
    
    def process_multiple_files(self, file_paths):
        """Process multiple files and combine OCR results"""
        all_text = ""
        processed_count = 0
        
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Processing Files")
        progress_window.geometry("400x150")
        
        progress_var = tk.StringVar()
        progress_label = ttk.Label(progress_window, textvariable=progress_var)
        progress_label.pack(pady=20)
        
        progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
        progress_bar.pack(pady=10)
        progress_bar['maximum'] = len(file_paths)
        
        def process_files():
            nonlocal all_text, processed_count
            
            for i, file_path in enumerate(file_paths):
                try:
                    progress_var.set(f"Processing: {os.path.basename(file_path)}")
                    progress_window.update()
                    
                    # Run OCR on each file
                    image = Image.open(file_path)
                    text = pytesseract.image_to_string(image)
                    
                    all_text += f"\n{'='*50}\n"
                    all_text += f"FILE: {os.path.basename(file_path)}\n"
                    all_text += f"{'='*50}\n"
                    all_text += text + "\n"
                    
                    processed_count += 1
                    progress_bar['value'] = i + 1
                    progress_window.update()
                    
                except Exception as e:
                    all_text += f"\nERROR processing {file_path}: {str(e)}\n"
                    continue
            
            # Display results
            self.text_display.delete(1.0, tk.END)
            self.text_display.insert(1.0, all_text)
            self.extracted_text = all_text
            
            progress_window.destroy()
            self.update_status(f"Batch processing complete: {processed_count}/{len(file_paths)} files processed")
        
        threading.Thread(target=process_files, daemon=True).start()
    
    def run_ocr(self):
        """Run OCR on the current image"""
        if self.current_image is None:
            messagebox.showwarning("Warning", "Please select an image first.")
            return
        
        self.update_status("Running OCR analysis...")
        
        def ocr_thread():
            try:
                # Run OCR
                self.extracted_text = pytesseract.image_to_string(self.current_image)
                
                # Update GUI in main thread
                self.root.after(0, lambda: self.display_ocr_results())
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("OCR Error", f"OCR failed: {str(e)}"))
                self.root.after(0, lambda: self.update_status("OCR failed"))
        
        threading.Thread(target=ocr_thread, daemon=True).start()
    
    def display_ocr_results(self):
        """Display OCR results in the text widget"""
        self.text_display.delete(1.0, tk.END)
        self.text_display.insert(1.0, self.extracted_text)
        self.update_status("OCR completed successfully")
    
    def save_as_text(self):
        """Save extracted text as .txt file"""
        if not self.extracted_text.strip():
            messagebox.showwarning("Warning", "No text to save. Run OCR first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Save as Text File"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.extracted_text)
                messagebox.showinfo("Success", f"Text saved to: {file_path}")
                self.update_status(f"Text saved: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")
    
    def save_as_pdf(self):
        """Save extracted text as PDF file"""
        if not self.extracted_text.strip():
            messagebox.showwarning("Warning", "No text to save. Run OCR first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            title="Save as PDF File"
        )
        
        if file_path:
            try:
                # Create PDF
                doc = SimpleDocTemplate(file_path, pagesize=A4)
                styles = getSampleStyleSheet()
                story = []
                
                # Add title
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=16,
                    spaceAfter=30,
                )
                story.append(Paragraph("OCR Extracted Text", title_style))
                story.append(Spacer(1, 12))
                
                # Add extracted text
                text_style = styles['Normal']
                text_lines = self.extracted_text.split('\n')
                
                for line in text_lines:
                    if line.strip():
                        story.append(Paragraph(line, text_style))
                    else:
                        story.append(Spacer(1, 6))
                
                # Build PDF
                doc.build(story)
                
                messagebox.showinfo("Success", f"PDF saved to: {file_path}")
                self.update_status(f"PDF saved: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")
    
    def clear_all(self):
        """Clear all content"""
        self.current_image = None
        self.extracted_text = ""
        self.text_display.delete(1.0, tk.END)
        self.image_canvas.delete("all")
        self.update_status("Ready - Select an input method to begin")

def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = OCRApplication(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()