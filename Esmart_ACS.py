from kensingtonapp import KensingtonApp  # Import KensingtonApp
import tkinter as tk
from tkinter import ttk  # ttk for modern styling
import subprocess  # To open Notepad

def open_newmen_app():
    """This function opens the main CheckingApp window."""
    import newmenapp
    root = tk.Tk()
    app = newmenapp.CheckingApp(root)
    root.mainloop()

def open_kensington_app():
    """This function opens the KensingtonApp window."""
    root = tk.Tk()
    app = KensingtonApp(root)
    root.mainloop()

def open_readme():
    """Open the readme file in Notepad."""
    subprocess.Popen(['notepad.exe', 'readme.txt'])  # Replace with the actual path

def Esmart_ACS():
    # Create the startup window
    startup_window = tk.Tk()
    startup_window.title("Esmart Checker System Launcher")

    # Set window size and background color
    startup_window.geometry("720x400")
    startup_window.configure(bg="#e8f0fe")  # Light blue background

    # Create a welcome label with a custom font and color
    welcome_label = ttk.Label(
        startup_window, 
        text="Esmart Auto Checking System", 
        font=("Helvetica", 30, "bold"), 
        foreground="#F87E05", 
        background="#e8f0fe"
    )
    welcome_label.pack(pady=20)

    # Configure ttk styles for the buttons
    style = ttk.Style()
    style.configure(
        "Custom.TButton", 
        font=("Helvetica", 14), 
        padding=10,
        relief="raised",
        background="#ffffff",  # White button background
        borderwidth=2
    )
    style.map("Custom.TButton", background=[('active', '#32CD32')])  # Change on hover

    # Button to start the NewmenApp (Internal Order Checker)
    start_newmen_button = ttk.Button(
        startup_window, 
        text="Internal Order Checker (Newmen) 新贵内部订单", 
        command=lambda: [startup_window.destroy(), open_newmen_app()],
        style="Custom.TButton"
    )
    start_newmen_button.pack(pady=10, padx=20, fill='x')  # Add padding and fill width

    # Button to start the KensingtonApp (Kensington Order Checker)
    start_kensington_button = ttk.Button(
        startup_window, 
        text="Work Order Checker (Kensington)", 
        command=lambda: [startup_window.destroy(), open_kensington_app()],
        style="Custom.TButton"
    )
    start_kensington_button.pack(pady=10, padx=20, fill='x')  # Add padding and fill width

     # Button to open the Readme / User Manual
    readme_button = ttk.Button(
        startup_window, 
        text="Readme / User Manual 使用手册", 
        command=open_readme,
        style="Custom.TButton"
    )
    readme_button.pack(pady=10, padx=20, fill='x')  # Add padding and fill width

    # Run the startup window
    startup_window.mainloop()

if __name__ == "__main__":
    Esmart_ACS()
