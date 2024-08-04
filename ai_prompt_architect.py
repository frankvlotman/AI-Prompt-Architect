import tkinter as tk
from tkinter import scrolledtext, messagebox
from langchain_openai import OpenAI
from langchain.prompts.prompt import PromptTemplate
from langchain_openai import ChatOpenAI
from langchain.chains import LLMChain
import os
from docx import Document
import win32com.client as win32
from PIL import Image  # Additional import for image handling

# Set the OpenAI API key
os.environ["OPENAI_API_KEY"] = "input API key here"

# Initialize the OpenAI model
llm = ChatOpenAI(model="gpt-4o-mini")

# Create a prompt template
prompt_template = PromptTemplate(
    input_variables=["prompt"],
    template="Prompt: {prompt}\nResponse:"
)

# Create the LLM chain
chain1 = prompt_template | llm

# Function to get response from OpenAI
def get_response():
    user_prompt = prompt_entry.get("1.0", tk.END).strip()
    if user_prompt:
        status_label.config(text="Fetching response...")
        root.update_idletasks()
        try:
            response = chain1.invoke({"prompt": user_prompt})
            content = response.content
            response_display.delete("1.0", tk.END)
            response_display.insert(tk.END, content)
            status_label.config(text="Response received")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            status_label.config(text="Error fetching response")
    else:
        messagebox.showwarning("Input Error", "Please enter a prompt.")
        status_label.config(text="")

# Function to export response to a .docx file
def export_response():
    content = response_display.get("1.0", tk.END).strip()
    if content:
        try:
            status_label.config(text="Exporting response to .docx...")
            root.update_idletasks()
            doc = Document()
            doc.add_heading('Hello', level=1)
            doc.add_paragraph(content)
            doc.save('C:\\Users\\Frank\\Desktop\\response.docx')
            status_label.config(text="Response exported to .docx")
            messagebox.showinfo("Export Success", "Response exported successfully.")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting: {e}")
            status_label.config(text="Error exporting response")

# Function to email the response using Outlook
def email_response():
    content = response_display.get("1.0", tk.END).strip()
    if content:
        try:
            status_label.config(text="Sending email...")
            root.update_idletasks()
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'email address here'
            mail.Subject = 'Hello'
            mail.Body = content
            mail.Send()
            status_label.config(text="Email sent")
            messagebox.showinfo("Email Sent", "Email sent successfully.")
        except Exception as e:
            messagebox.showerror("Email Error", f"An error occurred while sending email: {e}")
            status_label.config(text="Error sending email")

# Setting up the GUI
root = tk.Tk()
root.title("Hello")

# Prompt label and text entry
prompt_label = tk.Label(root, text="Enter your prompt:")
prompt_label.pack(pady=5)
prompt_entry = scrolledtext.ScrolledText(root, height=10, width=50)
prompt_entry.pack(pady=5)

# Response label and display
response_label = tk.Label(root, text="Response:")
response_label.pack(pady=5)
response_display = scrolledtext.ScrolledText(root, height=10, width=50)
response_display.pack(pady=5)

# Status label
status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=5)

# Frame to hold the buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

# Submit button
submit_button = tk.Button(button_frame, text="Submit", command=get_response)
submit_button.pack(side=tk.LEFT, padx=10)

# Export button
export_button = tk.Button(button_frame, text="Export", command=export_response)
export_button.pack(side=tk.LEFT, padx=10)

# Email button
email_button = tk.Button(button_frame, text="Email", command=email_response)
email_button.pack(side=tk.LEFT, padx=10)

# Start the GUI event loop
root.mainloop()
