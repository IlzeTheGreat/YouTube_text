import tkinter as tk
from tkinter import filedialog, messagebox
from youtube_transcript_api import YouTubeTranscriptApi
from docx import Document
import re
from datetime import datetime

def get_video_id(url):
    """
    Extract the video ID from a YouTube URL.
    """
    pattern = r"(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/(?:[^\/\n\s]+\/\S+\/|(?:v|e(?:mbed)?)\/|\S*?[?&]v=)|youtu\.be\/)([a-zA-Z0-9_-]{11})"
    match = re.match(pattern, url)
    return match.group(1) if match else None

def download_subtitles(video_id):
    """
    Download subtitles from a YouTube video.
    """
    transcript = YouTubeTranscriptApi.get_transcript(video_id)
    subtitles = ""
    for entry in transcript:
        subtitles += entry['text'] + " "
    return subtitles

def save_text_to_word(text, output_path):
    """
    Save text to a Word document.
    """
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_path)

def browse_directory():
    folder_selected = filedialog.askdirectory()
    save_path_entry.delete(0, tk.END)
    save_path_entry.insert(0, folder_selected)

def extract_and_save():
    youtube_url = url_entry.get()
    save_folder = save_path_entry.get()
    video_id = get_video_id(youtube_url)

    if not youtube_url or not save_folder:
        messagebox.showerror("Input Error", "Please provide both YouTube URL and save folder path.")
        return

    if video_id:
        try:
            subtitles_text = download_subtitles(video_id)
            today = datetime.today().strftime('%Y_%m_%d_%H_%M')
            file_name = f"{today}.docx"
            word_output_path = f"{save_folder}/{file_name}"
            save_text_to_word(subtitles_text, word_output_path)
            messagebox.showinfo("Success", f"Subtitles have been saved to {word_output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    else:
        messagebox.showerror("Input Error", "Invalid YouTube URL")

# Izveidot GUI
root = tk.Tk()
root.title("YouTube Subtitle Extractor")
root.configure(bg='#810C54')

# Pievienot ikonu logam
root.iconbitmap(r'C:\00_Ilze\ITIlze\Video_to_text\Video_camera.ico')  # Aizvietojiet ar jūsu ikonas faila nosaukumu

style = {
    'font': ('Helvetica', 12, 'bold'),
    'bg': '#810C54',
    'fg': 'white'
}

tk.Label(root, text="YouTube Video URL:", **style).grid(row=0, column=0, padx=10, pady=10)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Save Folder:", **style).grid(row=1, column=0, padx=10, pady=10)
save_path_entry = tk.Entry(root, width=50)
save_path_entry.grid(row=1, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="Browse", command=browse_directory, font=('Helvetica', 10, 'bold'), bg='#810C54', fg='white')
browse_button.grid(row=1, column=2, padx=10, pady=10)

extract_button = tk.Button(root, text="Extract and Save", command=extract_and_save, font=('Helvetica', 10, 'bold'), bg='#810C54', fg='white')
extract_button.grid(row=2, column=0, columnspan=3, padx=10, pady=20)

root.mainloop()
