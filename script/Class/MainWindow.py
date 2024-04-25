import threading
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

from Class_UploadMediaPresentation import UploadMediaPresentation


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Upload Media to PowerPoint")
        self.geometry("400x200")

        self.media_folder_label = tk.Label(self, text="Media Folder:")
        self.media_folder_label.pack()
        self.media_folder_entry = tk.Entry(self, width=50)
        self.media_folder_entry.pack()
        self.media_folder_button = tk.Button(
            self, text="Browse", command=self.browse_media_folder
        )
        self.media_folder_button.pack()

        self.pptx_path_label = tk.Label(self, text="PowerPoint File:")
        self.pptx_path_label.pack()
        self.pptx_path_entry = tk.Entry(self, width=50)
        self.pptx_path_entry.pack()
        self.pptx_path_button = tk.Button(
            self, text="Browse", command=self.browse_pptx_file
        )
        self.pptx_path_button.pack()

        self.start_button = tk.Button(self, text="Start", command=self.start_upload)
        self.start_button.pack()

    def browse_media_folder(self):
        folder_path = filedialog.askdirectory()
        self.media_folder_entry.delete(0, tk.END)
        self.media_folder_entry.insert(0, folder_path)

    def browse_pptx_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PowerPoint Files", "*.pptx")]
        )
        self.pptx_path_entry.delete(0, tk.END)
        self.pptx_path_entry.insert(0, file_path)

    def start_upload(self):
        media_folder = self.media_folder_entry.get()
        pptx_path = self.pptx_path_entry.get()

        if not media_folder or not pptx_path:
            messagebox.showerror(
                "Error", "Please provide both Media Folder and PowerPoint File paths."
            )
            return

        media_presentation = UploadMediaPresentation(
            media_folder=media_folder, pptx_path=pptx_path
        )
        thread = threading.Thread(target=media_presentation.Start)
        thread.start()
        messagebox.showinfo(
            "Success", "Media upload started. Please wait for completion."
        )


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
