import logging
import os
import threading
from pptx import Presentation

VIDEO_FOLDER = "folder1"
PPTX_PATH = "Java-folder1.pptx"
FILE_LOG = "app.log"
logging.basicConfig(filename=FILE_LOG, level=logging.INFO)

prs = Presentation()


# לולאה על כל קובץ בתיקיית הוידאו
def upload_Video_To_PPTX(video_name):
    print(f"in {FILE_LOG}Record all events")
    for filename in os.listdir(video_name):
        if filename.endswith(".mp4"):
            video_path = os.path.join(video_name, filename)
            video_name_without_extension = os.path.splitext(filename)[0]
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title = slide.shapes.title
            title.text = video_name_without_extension
            try:
                slide.shapes.add_movie(
                    video_path,
                    left=0,
                    top=0,
                    width=prs.slide_width,
                    height=prs.slide_height,
                )
                logging.info(f"Uploaded file: {video_name_without_extension}")
            except Exception as err:

                logging.error(
                    f"Error uploading file {video_name_without_extension}: {err}"
                )
    else:
        logging.NOTSET(f"success")


# שמירת המצגת
def File_Is_Existed(PPTX_PATH):
    if os.path.exists(PPTX_PATH):
        print("Adding to an existing PowerPoint file:{PPTX_PATH}")
    else:
        print(f"cerate new PowerPoint file:\t{PPTX_PATH}")


if __name__ == "__main__":
    File_Is_Existed(PPTX_PATH)
    try:
        thread = threading.Thread(target=upload_Video_To_PPTX, args=(VIDEO_FOLDER,))
        thread.start()
        thread.join()
        prs.save(PPTX_PATH)
    except Exception as e:
        print(f"Error: {e}")
        logging.error(f"Error: {e}")
