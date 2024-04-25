# Upload Media to PowerPoint

This Python project allows you to upload media files (images, videos, and audio) from a folder to a PowerPoint presentation. It creates a new PowerPoint file or appends to an existing one, with each media file placed on a separate slide.

## Requirements

- Python 3.x
- python-pptx library (install using `pip install python-pptx`)

## Usage

1. Run the `main.py` script.
2. In the GUI window, click the "Browse" button next to "Media Folder" to select the folder containing the media files you want to upload.
3. Click the "Browse" button next to "PowerPoint File" to select an existing PowerPoint file or specify a new file name with a `.pptx` extension.
4. Click the "Start" button to begin uploading the media files to the PowerPoint presentation.
5. Wait for the process to complete. A message box will appear when the upload is finished.

## Project Structure

- `main.py`: The main script that runs the application.
- `Class_UploadMediaPresentation.py`: Contains the `UploadMediaPresentation` class responsible for uploading media files to the PowerPoint presentation.
- `MainWindow.py`: Defines the `MainWindow` class, which creates the GUI for the application using Tkinter.

## Key Features

- Supports uploading images (JPG, PNG, GIF), videos (MP4), and audio files (WAV, MP3) to PowerPoint.
- Creates a new slide for each media file, with the file name as the slide title.
- Media files are scaled to fit the slide dimensions.
- Logging functionality to track the upload process and any errors that may occur.
- GUI interface for easy selection of media folder and PowerPoint file paths.

## Customization

You can customize the project by modifying the following aspects:

- Presentation title: Change the value of the `presentation_title` parameter in the `UploadMediaPresentation` class constructor.
- Log file path: Change the value of the `log_file` parameter in the `UploadMediaPresentation` class constructor.
- Slide layout: Modify the `self.prs.slide_layouts[5]` line in the `upload_media_to_pptx` method to use a different slide layout.
- GUI appearance: Modify the `MainWindow` class in `MainWindow.py` to change the appearance and layout of the GUI.

Feel free to explore and modify the code to suit your specific needs.
