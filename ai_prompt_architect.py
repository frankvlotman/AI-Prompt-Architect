import tkinter as tk
from tkinter import scrolledtext, messagebox, Toplevel, Label, Text, Button
from langchain_openai import ChatOpenAI
from langchain.prompts.prompt import PromptTemplate
from langchain.chains import LLMChain
import os
import win32com.client as win32
from docx import Document
import pyttsx3

# Set the OpenAI API key #
os.environ["OPENAI_API_KEY"] = "input API key here"

# Initialize the OpenAI model
llm = ChatOpenAI(model_name="gpt-4o-mini")

# Create a prompt template
prompt_template = PromptTemplate(
    input_variables=["prompt"],
    template="Prompt: {prompt}\nResponse:"
)

# Create the LLM chain
chain1 = LLMChain(prompt=prompt_template, llm=llm)

# Function to get response from OpenAI and display it in a new window #####
def get_response_and_display():
    user_prompt = prompt_entry.get("1.0", tk.END).strip()
    if user_prompt:
        status_label.config(text="Fetching response...")
        root.update_idletasks()
        try:
            response = chain1({"prompt": user_prompt})
            content = response["text"]
            status_label.config(text="Response received, displaying...")
            root.update_idletasks()
            
            # Create a new window for displaying the response
            response_window = Toplevel(root)
            response_window.title("Response")
            response_window.geometry("600x600")  # Increased height to fit buttons without resizing         

            # Display the response in the new window
            response_label = Label(response_window, text="Response:", font=("Arial", 14))
            response_label.pack(pady=10)
            
            response_text = Text(response_window, wrap=tk.WORD, font=("Arial", 12))
            response_text.insert(tk.END, content)
            response_text.config(state=tk.DISABLED)
            response_text.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)
            
            def download_to_docx():
                filepath = "C:\\Users\\Frank\\Desktop\\response.docx"
                doc = Document()
                doc.add_paragraph(content)
                doc.save(filepath)
                messagebox.showinfo("Saved", f"Response saved to {filepath}")
            
            def send_email():
                try:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = 'email@email.com'
                    mail.Subject = 'Hello'
                    mail.Body = content
                    mail.Send()
                    messagebox.showinfo("Email Sent", "Email sent successfully.")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred while sending email: {e}")

            def play_text():
                engine = pyttsx3.init()
                
                # Get the list of available voices
                voices = engine.getProperty('voices')
                
                # Select the desired voice (e.g., female voice)
                for voice in voices:
                    if "female" in voice.name.lower():
                        engine.setProperty('voice', voice.id)
                        break  # Break the loop once a female voice is found
                
                # Optionally, you can set the speech rate (words per minute)
                engine.setProperty('rate', 150)  # Adjust the rate if needed

                # Optionally, you can set the volume (0.0 to 1.0)
                engine.setProperty('volume', 1.0)  # Set volume to maximum

                engine.say(content)
                engine.runAndWait()
            
            # Buttons for downloading, sending email, and playing text
            button_frame = tk.Frame(response_window)
            button_frame.pack(pady=10)

            download_button = Button(button_frame, text="Download", command=download_to_docx, bg='#d0e8f1')
            download_button.pack(side=tk.LEFT, padx=10)

            email_button = Button(button_frame, text="Email", command=send_email, bg='#d0e8f1')
            email_button.pack(side=tk.LEFT, padx=10)

            tts_button = Button(button_frame, text="Play Text", command=play_text, bg='#d0e8f1')
            tts_button.pack(side=tk.RIGHT, padx=10)
            
            status_label.config(text="Response displayed")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            status_label.config(text="Error")
    else:
        messagebox.showwarning("Input Error", "Please enter a prompt.")
        status_label.config(text="")

# Setting up the GUI
root = tk.Tk()
root.title("Hello")

# Prompt label and text entry
prompt_label = tk.Label(root, text="Enter your prompt:")
prompt_label.pack(pady=5)
prompt_entry = scrolledtext.ScrolledText(root, height=10, width=50)
prompt_entry.pack(pady=5)

# Status label
status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=5)

# Frame to hold the buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

# Submit button
submit_button = tk.Button(button_frame, text="Submit", command=get_response_and_display, bg='#d0e8f1')
submit_button.pack(side=tk.LEFT, padx=10)

# Start the GUI event loop
root.mainloop()
