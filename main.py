# PROFESSIONAL PROJECT: Convert PDF to Audiobook

# OBJECTIVE:  To extract contents from a PDF file and convert them to audio,
# thus creating an audiobook.

# Import necessary library(ies):
import boto3 # Amazon Web Services (AWS) Software Development Kit (SDK) for Python
from datetime import datetime
from dotenv import load_dotenv
import math
from moviepy.editor import concatenate_audioclips, AudioFileClip  # Used to combine audio files when text length > 3000 characters.
import os
from pypdf import PdfReader
from tkinter import *
from tkinter import messagebox, filedialog
import traceback
import wx
import wx.lib.agw.pybusyinfo as PBI

# Load environmental variables from the ".env" file:
load_dotenv()

# Define constants for storing environmental variables:
AMAZON_POLLY_ACCESS_KEY = None
AMAZON_POLLY_SECRET_ACCESS_KEY = None

# Define variable to be used for showing in-process updates (used in the text-to-audio conversion process):
dlg = wx.App()

# Define constants for application default font size as well as window's height and width:
FONT_NAME = "Arial"
WINDOW_HEIGHT = 300
WINDOW_WIDTH = 350

# Define constant for API to use in conversion:
AMAZON_POLLY_API = "https://polly.us-east-1.amazonaws.com/v1/speech/"

# Define variable for the GUI (application) window (so that it can be used globally), and make it a Tkinter instance:
window = Tk()

# Define variable to store image on main page of application:
img = None

# Define variable for storing path/filename for file selected for conversion:
selected_file = ""


# DEFINE FUNCTIONS TO BE USED FOR THIS APPLICATION (LISTED IN ALPHABETICAL ORDER BY FUNCTION NAME):
def handle_window_on_closing():
    """Function which confirms with user if s/he wishes to exit this application"""
    try:
        # Confirm with user if s/he wishes to exit this application.  If confirmed, exit:
        if messagebox.askokcancel("Exit?", "Do you want to exit this application?"):
            exit()
    except:
        # Exit the application:
        exit()


def read_and_convert_file():
    """Function for enabling user to select a PDF file to convert to audio"""
    global selected_file, dlg

    try:
        # Open a file dialog window so that user can select file for conversion.  Store the selected file's path/filename
        # into a variable for further reference/use:
        selected_file = filedialog.askopenfilename(initialdir=os.getcwd(), filetypes=[("PDF files", "*.pdf")])

        if selected_file != "":  # User has selected a file to convert.
            # Identify the path/name of the file which will store the final (combined where total text > 3000) audio contents:
            output_file_final = selected_file[0:selected_file.find(".pdf")] + ".mp3"

            # Create a pdf reader object:
            reader = PdfReader(selected_file)

            # Extract text from selected file:
            text_to_convert = ""
            for pg in reader.pages:
                text_to_convert += pg.extract_text()

            # Get the total length of the text to be converted:
            length_text = len(text_to_convert)

            # Get the total # of portions needed to span the entirety of the total length of content:
            num_portions = math.ceil(length_text / 3000)

            # Configure the Polly client:
            polly = boto3.client("polly",
                                region_name="us-east-1",
                                aws_access_key_id=AMAZON_POLLY_ACCESS_KEY,
                                aws_secret_access_key=AMAZON_POLLY_SECRET_ACCESS_KEY)

            # Initialize a list variable to store names of files produces in conversion
            # to audio.  This is particularly necessary for source files where text
            # length > 3000 and, therefore, warranting processing in 3000-character-or-less
            # portions until entirety of source file's text has been accounted for.:
            converted_files = []

            # Initialize variables to be used in processing source file in portions as
            # described earlier:
            index_text_start = 0  # Start of slicing
            index_text_end = 3000  # End of slicing
            file_in_process = 1  # Number of files in process of conversion

            # Initiate a variable to support dialog for keeping user informed on update status:
            app_wx = wx.App(redirect=False)

            # Loop through the source file and convert it to audio in 3000-character-or-less
            # portions until entirety of source file's text has been accounted for:
            while file_in_process <= num_portions:
                dlg = PBI.PyBusyInfo(f"Processing portion {file_in_process} of {num_portions}...", title="Conversion to audio")
                if index_text_start > length_text:
                    break

                # If less than 3000 characters remain to be processed, make the ending "slicing"
                # variable to equal the length of the source file's entire text:
                if index_text_start + 3000 > length_text:
                    index_text_end = length_text

                # Convert text to speech using the "Polly boto3" API client:
                spoken_text = polly.synthesize_speech(Text=text_to_convert[index_text_start:index_text_end],
                                                      OutputFormat='mp3', VoiceId='Joanna')

                # Identify the path/name of file to which the converted audio should be stored:
                output_file = selected_file[0:selected_file.find(".pdf")] + "_" + ("000" + str(file_in_process))[-4:] + ".mp3"

                # Write spoken text (speech) to an audio file (format = mp3):
                with open(output_file, 'wb') as f:
                    f.write(spoken_text['AudioStream'].read())
                    f.close()

                # Append path/filename of converted (audio) file to the "converted_files" list variable:
                converted_files.append(output_file)

                # Increment "slicing" variables by 3000 and the file tracker by 1:
                index_text_start += 3000
                index_text_end += 3000
                file_in_process += 1

                # Reinitialize the in-process dialog box variable:
                dlg = wx.App()

            # If final output file already exists, delete it prior to the next step:
            if os.path.exists(output_file_final):
                dlg = PBI.PyBusyInfo(f"Removing {output_file_final}...", title="Conversion to audio")
                os.remove(output_file_final)
                dlg = wx.App()

            # If file was processed in 3000-character portions and total text length is > 3000, combine all
            # portion files into one.  Otherwise, rename lone "portion" file to match path/name of source file
            # (albeit with ".mp3" extension):
            dlg = PBI.PyBusyInfo(f"Combining portions into one audio file ({output_file_final})...", title="Conversion to audio")
            if len(converted_files) == 1:
                os.rename(converted_files[0], output_file_final)
            else:
                # Combine "portion" files into one and store it under the final output path/filename:
                audio_portions = [AudioFileClip(file) for file in converted_files]
                final_audio = concatenate_audioclips(audio_portions)
                final_audio.write_audiofile(output_file_final)

                # Delete "portion" files that were created via the process above:
                for file in converted_files:
                    dlg = wx.App()
                    dlg = PBI.PyBusyInfo(f"Deleting file {file}...", title="Conversion to audio")
                    if os.path.exists(file):
                        os.remove(file)

            # Reinitialize dialog variable:
            dlg = wx.App()

            # Destroy dialog app:
            app_wx.MainLoop()

            # Inform user that file conversion has been completed.
            messagebox.showinfo("File Conversion Completed",
                                f"File has been converted to audio and stored at:\n\n{output_file_final}")


    except:
        messagebox.showinfo("Error", f"Error (read_and_convert_file): {traceback.format_exc()}")
        update_system_log("read_and_convert_file", traceback.format_exc())
        dlg = wx.App()


def run_app():
    """Main function used to run this application"""
    global AMAZON_POLLY_ACCESS_KEY, AMAZON_POLLY_SECRET_ACCESS_KEY

    try:
        # Define constants for storing environmental variables:
        AMAZON_POLLY_ACCESS_KEY = os.getenv("AMAZON_POLLY_ACCESS_KEY")
        AMAZON_POLLY_SECRET_ACCESS_KEY = os.getenv("AMAZON_POLLY_SECRET_ACCESS_KEY")

        # Create and configure all visible aspects of the application window. If function did not
        # execute successfully, exit this application:
        if not window_config():
            exit()

        # From this point, subsequent functionality defined via the "read_and_convert_file" function (launched via the GUI).

    except:
        messagebox.showinfo("Error", f"Error (run_app): {traceback.format_exc()}")
        update_system_log("run_app", traceback.format_exc())
        exit()


def update_system_log(activity, log):
    """Function to update the system log with errors encountered"""
    try:
        # Capture current date/time:
        current_date_time = datetime.now()
        current_date_time_file = current_date_time.strftime("%Y-%m-%d")

        # Update log file.  If log file does not exist, create it:
        with open("log_convert_pdf_to_audiobook_" + current_date_time_file + ".txt", "a") as f:
            f.write(datetime.now().strftime("%Y-%m-%d @ %I:%M %p") + ":\n")
            f.write(activity + ": " + log + "\n")

        # Close the log file:
        f.close()

    except:
        messagebox.showinfo("Error", f"Error: System log could not be updated.\n{traceback.format_exc()}")


def window_center_screen():
    """Function which centers the application window on the computer screen"""
    global window

    try:
        # Capture the desired width and height for the window:
        w = WINDOW_WIDTH # width of tkinter window
        h = WINDOW_HEIGHT  # height of tkinter window

        # Capture the computer screen's width and height:
        screen_width = window.winfo_screenwidth()  # Width of the screen
        screen_height = window.winfo_screenheight()  # Height of the screen

        # Calculate starting X and Y coordinates for the application window:
        x = (screen_width / 2) - (w / 2)
        y = (screen_height / 2) - (h / 2)

        # Center the application window based on the aforementioned constructs:
        window.geometry('%dx%d+%d+%d' % (w, h, x, y))

        # At this point, function is presumed to have executed successfully.
        # Return successful-execution indication to the calling function:
        return True

    except:
        messagebox.showinfo("Error",f"Error (window_center_screen): {traceback.format_exc()}")
        update_system_log("window_center_screen", traceback.format_exc())
        return False


def window_config():
    """Function which creates and configures all visible aspects of the application window"""
    try:
        # Create and configure application window.  If function did not execute successfully,
        # return failed-execution indication to the calling function:
        if not window_create_and_config():
            return False

        # Create and configure application's user interface. If function did not execute successfully,
        # return failed-execution indication to the calling function:
        if not window_create_and_config_user_interface():
            return False

        # At this point, function is presumed to have executed successfully.
        # Return successful-execution indication to the calling function:
        return True

    except:
        messagebox.showinfo("Error", f"Error (window_config): {traceback.format_exc()}")
        update_system_log("window_config", traceback.format_exc())
        return False


def window_create_and_config():
    """Function to create and configure the GUI (application) window"""
    global window

    try:
        # Create and configure the application window:
        window.title("Convert PDF to Audiobook")
        window.minsize(width=WINDOW_WIDTH, height=WINDOW_HEIGHT)
        window.config(padx=25, pady=0,bg='teal')
        window.resizable(0, 0)  # Prevents window from being resized.
        window.attributes("-toolwindow", 1)  # Removes the minimize and maximize buttons from the application window.
        window.attributes('-topmost', True)  # Keeps application's main window on top.

        # Center the application window on the computer screen. If function did not
        # execute successfully, return failed-execution indication to the calling function::
        if not window_center_screen():
            return False

        # Prepare the application to handle the event of user attempting to close the application window:
        window.protocol("WM_DELETE_WINDOW", handle_window_on_closing)

        # At this point, function is presumed to have executed successfully.
        # Return successful-execution indication to the calling function:
        return True

    except:
        messagebox.showinfo("Error", f"Error (window_create_and_config): {traceback.format_exc()}")
        update_system_log("window_create_and_config", traceback.format_exc())
        return False


def window_create_and_config_user_interface():
    """Function which creates and configures items comprising the user interface, including the canvas (which overlays on top of the app. window), labels, and button"""
    global img, window

    try:
        # Create and configure canvas which overlays on top of window:
        canvas = Canvas(window)
        img = PhotoImage(file="audiobook.png")
        canvas.config(height=img.height(), width=img.width(), bg='teal', highlightthickness=0)
        canvas.create_image(100,60, image=img)
        canvas.grid(column=0, row=1, columnspan=3, padx=0, pady=0)
        canvas.update()

        # Create and configure the introductory header text (label):
        label_intro = Label(text=f"WELCOME TO MY\nPDF-TO-AUDIO CONVERTER!", height=3, bg='teal', fg='black', padx=0, pady=0, font=(FONT_NAME,16, "bold"))
        label_intro.grid(column=0, row=0, columnspan=3)

        # Create and configure the blank "label" which serves as a separator between the canvas image and the file-selection button:
        label_space = Label(text="", bg='teal', fg='white', padx=0, pady=0, font=(FONT_NAME,16, "bold"))
        label_space.grid(column=0, row=2)

        # Create and configure button used to select file for conversion:
        button_select_file = Button(text="Select File To Convert", width=20, height=1, bg='grey', fg='white', pady=0, font=(FONT_NAME,14,"bold"), command=read_and_convert_file)
        button_select_file.grid(column=0, row=3,columnspan=6)

        # At this point, function is presumed to have executed successfully.
        # Return successful-execution indication to the calling function:
        return True

    except:
        messagebox.showinfo("Error", f"Error (window_create_and_config_user_interface): {traceback.format_exc()}")
        update_system_log("window_create_and_config_user_interface", traceback.format_exc())
        return False


# Run the application:
run_app()

# Keep application window open until user closes it:
window.mainloop()

if __name__ == '__main__':
    run_app()
