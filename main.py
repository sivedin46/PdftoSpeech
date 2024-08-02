from PyPDF2 import PdfReader
import pyttsx3
from tkinter import *
from tkinter import filedialog, messagebox
import threading
import importlib
import ffmpeg
import os
import shutil


class PdfSpeech:
    def __init__(self):
        self.create_window()
        self.create_ui_components()
        self.first_initialize()
        self.window.mainloop()

    def first_initialize(self):
        self.word_read_button.config(state="disabled")
        self.direct_read_button.config(state="disabled")
        self.pause_button.config(state="disabled")
        self.stop_button.config(state="disabled")
        self.engine = pyttsx3.init(driverName='sapi5')
        self.voices = self.engine.getProperty('voices')
        self.speaker = self.voices[0].id
        self.load_speakers()
        self.speed = 125
        self.volume_ = 0.5
        self.engine.setProperty('volume', self.volume_)
        self.engine.setProperty('rate', self.speed)
        self.number_of_pages = 0
        self.char_counter = 0
        self.char_index = 0
        self.word_index = 0
        self.mode = 0
        self.is_paused = False
        self.is_stopped = False  # Add this line
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.lock = threading.Lock()

    def create_window(self):
        self.window = Tk()
        self.window.title("PDF TO SPEECH - @lper")
        self.window.config(padx=20, pady=20, bg="salmon3")
        self.window.geometry("800x400")
        self.window.resizable(None)
        self.i = PhotoImage(width=1, height=1)

    def create_ui_components(self):
        self.create_info_frame()
        self.create_text_frame()
        self.create_entry_frame()
        self.create_button_frame()
        self.create_slider_frame()

    def create_info_frame(self):
        info_frame = Frame(self.window, bg="salmon3", highlightthickness=0)
        info_frame.pack(side="top", fill="x", pady=3)
        self.pages_info_label = Label(info_frame, text="Total Pages: ", bg="salmon3")
        self.pages_info_label.pack(side=LEFT, padx=5)

    def create_text_frame(self):
        text_frame = Frame(self.window, bg="salmon3", highlightthickness=0)
        text_frame.pack(side="top", fill="x", pady=3)
        text_scrollbar = Scrollbar(text_frame, orient=VERTICAL)
        text_scrollbar.pack(side="right", fill=Y)
        self.text_entry = Text(text_frame, width=590, height=10, font=("Arial", "14"), bg="white", wrap=WORD,
                               yscrollcommand=text_scrollbar.set)
        self.text_entry.insert(END, "Please load a PDF file for PDF-Speech exercise. \n")
        self.text_entry.propagate(False)
        self.text_entry.tag_configure("highlight", foreground="green")
        self.text_entry.pack(side="left", padx=5, fill="both", expand=True)
        text_scrollbar.config(command=self.text_entry.yview)

    def create_entry_frame(self):
        entry_frame = Frame(self.window, bg="salmon3", highlightthickness=0)
        entry_frame.pack(side="top", fill="x", pady=3)

    def create_button_frame(self):
        button_frame = Frame(self.window, bg="salmon3", highlightthickness=0)
        button_frame.pack(side="bottom", fill="x", pady=3)
        file_open_label = Label(button_frame, text="File Path:", bg="salmon3")
        file_open_label.pack(side="left", padx=5)
        self.file_entry = Entry(button_frame, width=35)
        self.file_entry.pack(side="left", padx=5)
        self.open_file_button = Button(button_frame, width=8, text="Open PDF", command=self.load_pdf_file)
        self.open_file_button.pack(side="left", padx=5)
        self.word_read_button = Button(button_frame, width=10, text="Read Words", command=self.word_reader_starter)
        self.word_read_button.pack(side="left", padx=5)
        self.direct_read_button = Button(button_frame, width=8, text="Read", command=self.sentence_reader_starter)
        self.direct_read_button.pack(side="left", padx=5)
        self.pause_button = Button(button_frame, width=8, text="Pause", command=self.pause_resume)
        self.pause_button.pack(side="left", padx=5)
        self.stop_button = Button(button_frame, width=8, text="Stop", command=self.stop_reading)
        self.stop_button.pack(side="left", padx=5)
        record_button = Button(button_frame, width=8, text="REC", bg="red", command=self.convert_to_wav)
        record_button.pack(side="left", padx=5)

    def create_slider_frame(self):
        slider_frame = Frame(self.window, bg="salmon3", highlightthickness=0)
        slider_frame.pack(side="bottom", fill="x", pady=3)
        language_label = Label(slider_frame, text="Language", bg="salmon3")
        language_label.pack(side="left", padx=5)
        scrollbar = Scrollbar(slider_frame, orient=VERTICAL)
        self.language_list = Listbox(slider_frame, width=30, height=10, selectmode="single", bg="white",
                                     highlightthickness=0, yscrollcommand=scrollbar.set)
        self.language_list.pack(side="left", padx=5)
        scrollbar.pack(side="left", fill=Y)
        scrollbar.config(command=self.language_list.yview)
        self.language_list.bind('<<ListboxSelect>>', self.select_speaker)
        volume_label = Label(slider_frame, text="Volume", bg="salmon3")
        volume_label.pack(side="left", padx=5)
        volume_value = IntVar(value=50)
        volume_button = Scale(slider_frame, width=20, length=100, orient="horizontal", from_=0, to=100, bg="salmon3",
                              highlightthickness=0, variable=volume_value)
        volume_button.config(command=self.adjust_volume)
        volume_button.pack(side="left", padx=5)
        speed_button_label = Label(slider_frame, text="Speech Speed", bg="salmon3")
        speed_button_label.pack(side="left", padx=5)
        speed_value = IntVar(value=125)
        speed_button = Scale(slider_frame, width=20, length=100, orient="horizontal", from_=50, to=300, bg="salmon3",
                             highlightthickness=0, variable=speed_value)
        speed_button.config(command=self.adjust_speed)
        speed_button.pack(side="left", padx=5)

    def reset_engine(self): # unfortunately pytts3 engine sometimes crashes and needs restart. after stop using it
        importlib.reload(pyttsx3)
        self.engine = pyttsx3.init(driverName='sapi5')
        self.engine.setProperty('voice', self.speaker)
        self.engine.setProperty('volume', self.volume_)
        self.engine.setProperty('rate', self.speed)

    def reset_variables(self): # resetting variables after stop button usage for starting reading or new pdf upload
        self.char_counter = 0
        self.mode = 0
        self.word_index = 0
        self.char_index = 0

    def load_pdf_file(self): # loads pdf file and creates words list  and text string for reading words or sentences
        self.text = ""
        self.words = []
        file_path = filedialog.askopenfilename(title="Open PDF File", filetypes=[("PDF", "*.pdf")])
        if not file_path:
            return
        self.file_entry.delete(0, END)
        self.file_entry.insert(index=0, string=file_path)
        try:
            reader = PdfReader(file_path)
            self.number_of_pages = len(reader.pages)
            self.pages_info_label.config(text=f"Total Pages: {self.number_of_pages}")
            for index in range(self.number_of_pages):
                page = reader.pages[index]
                self.text += page.extract_text()
                self.words = self.text.split()
            self.text_entry.delete(1.0, END)
            self.text_entry.insert(END, ' '.join(self.words))
            self.word_read_button.config(state="normal")
            self.direct_read_button.config(state="normal")
            self.pause_button.config(state="normal")
            self.stop_button.config(state="normal")
        except Exception as e:
            messagebox.showerror("Warning", f"An error occured while reading PDF File: {e}")

    def load_speakers(self):
        for index, voice in enumerate(self.voices): # detects windows speech lang database in shows user
            names = voice.name
            voice_option = names.split(" - ")[0].replace("Desktop", "").replace("Microsoft", "").strip()
            language = names.split(" - ")[1].replace("(", "").replace(")", "")
            sounds = f"{language} - {voice_option}"
            self.language_list.insert(END, f"{sounds} \n")

    def select_speaker(self, event): # detecs user choice and sets as voice for pytts3
        selected_indice = self.language_list.curselection()
        if selected_indice:
            speaker_index = selected_indice[0]
            self.speaker = self.voices[speaker_index].id
            self.engine.setProperty('voice', self.speaker)

    def adjust_volume(self, val):
        self.volume_ = float(val) / 100
        self.engine.setProperty('volume', self.volume_)  # setting up volume level  between 0 and 1

    def adjust_speed(self, val):
        self.speed = float(val)
        self.engine.setProperty('rate', float(self.speed))  # setting up new voice rate

    def read_words(self): # reads pdf word by word
        while self.word_index < len(self.words):
            if self.is_stopped:
                break
            if self.pause_event.is_set():
                word = self.words[self.word_index]
                chars = len(word) + 1
                start_index = f"1.{self.char_counter}"
                end_index = f"1.{chars + self.char_counter}"
                self.char_counter += chars
                self.text_entry.tag_add("highlight", start_index, end_index)
                with self.lock:  # locking everything while pytts3 engine is working.
                    self.engine.say(word)   # otherwise it changes variables or problems occur
                    self.engine.runAndWait()
                self.word_index += 1
            else:
                self.engine.stop()
                break
        if self.word_index >= len(self.words):
            self.stop_reading()
            messagebox.showinfo("Warning", f"Reading file completed")

    def read_sentences(self): # reads sentence(it can be set by buffer_size) detects if the pdf finished
        while self.char_index < len(self.text):
            if self.is_stopped:
                break
            if self.pause_event.is_set():
                buffer_size = 75
                end_index = min(self.char_index + buffer_size, len(self.text))
                buffer = self.text[self.char_index:end_index]
                if end_index < len(self.text) and buffer[-1] != " ":
                    while end_index < len(self.text) and self.text[end_index] != " ":
                        end_index += 1
                    buffer = self.text[self.char_index:end_index]
                start_index = f"1.{self.char_index}"
                end_index_str = f"1.{end_index}"
                self.text_entry.tag_add("highlight", start_index, end_index_str)
                self.char_index = end_index
                with self.lock:
                    self.engine.say(buffer)
                    self.engine.runAndWait()
            else:
                self.engine.stop()
                break
        if self.char_index >= len(self.text):
            self.stop_reading()
            messagebox.showinfo("Warning", f"Reading file completed")

    def reader(self, reader_funct): # controls the reading event depending on which button clicked
        self.is_stopped = False
        self.pause_event.set()
        self.word_read_button.config(state="disabled")
        self.open_file_button.config(state="disabled")
        self.direct_read_button.config(state="disabled")
        self.mode = 0 if reader_funct == self.read_words else 1
        if self.is_paused:
            self.is_paused = False
            self.pause_event.set()
        threading.Thread(target=reader_funct, daemon=True).start()

    def word_reader_starter(self): # word read button clicked
        self.reader(self.read_words)

    def sentence_reader_starter(self):# sentence button clicked
        self.reader(self.read_sentences)

    def pause_resume(self): #toggle between pause and resume depending on which mode is running word or sentence reading
        if self.is_paused:
            self.pause_event.set()
            self.is_paused = False
            self.pause_button.config(text="Pause")
            if self.mode == 0:
                threading.Thread(target=self.read_words, daemon=True).start()
            else:
                threading.Thread(target=self.read_sentences, daemon=True).start()
        else:
            self.pause_event.clear()
            self.is_paused = True
            self.pause_button.config(text="Resume")

    def stop_reading(self): # stops readings and resets engine and variables
        self.is_stopped = True
        self.pause_event.clear()
        self.engine.stop()
        self.is_paused = False
        self.pause_button.config(text="Pause")
        self.word_read_button.config(state="normal")
        self.direct_read_button.config(state="normal")
        self.open_file_button.config(state="normal")
        self.text_entry.tag_remove("highlight", "1.0", END)
        self.reset_engine()
        self.reset_variables()

    def convert_to_wav(self): # converts text to mp3
        temp_mp3 = "temp.mp3"
        self.engine.save_to_file(self.text, temp_mp3)
        self.engine.runAndWait()
        file = filedialog.asksaveasfile(mode='w', defaultextension=".mp3",
                                        filetypes=(("MP3 files", "*.mp3"), ("All Files", "*.*")))
        if file:
            abs_path = os.path.abspath(file.name)
            shutil.move(temp_mp3, abs_path)
            messagebox.showinfo("Warning", f"MP3 file saved as {abs_path}")
        else:
            os.remove(temp_mp3)
            messagebox.showinfo("Warning", "Save operation cancelled, temporary file removed.")


if __name__ == "__main__":
    PdfSpeech()
