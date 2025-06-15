import tkinter as tk
import customtkinter as ctk
from YFinanceModule2 import search_screener
import threading
import os
import sys
from BetaCalculatorModule import *
from datetime import datetime


if getattr(sys, 'frozen', False):
    # Running as a bundled executable
    # The _MEI_ path is where PyInstaller unpacks the files
    base_path = sys._MEIPASS # Use _MEI_TEMP if you use --temp-dir
    # If not using --temp-dir, use sys._MEI_
    # base_path = sys._MEI_
else:
    # Running as a regular Python script
    base_path = os.path.dirname(os.path.abspath(__file__))

download_path = os.path.join(base_path, "downloads")


instance = search_screener(download_path)

title = "Protocol 1 (Version 1.0)"

# UI Styling Constants
DARK_BG = "#1E1E1E"
LIGHT_BG = "#2D2D2D"
ACCENT_COLOR = "#3498db"  # Blue accent
HIGHLIGHT_COLOR = "#2980b9"  # Darker blue for highlights
TEXT_COLOR = "white"
SECONDARY_TEXT = "#CCCCCC"
BORDER_RADIUS = 10
FONT_MAIN = ("Segoe UI", 12)
FONT_HEADING = ("Segoe UI", 18, "bold")
FONT_SUBHEADING = ("Segoe UI", 14)
FONT_SMALL = ("Segoe UI", 10)

# Initialize app theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

window = ctk.CTk()
window.title(title)
window.geometry("700x700")
window.resizable(False, False)

# Global variable to track loading state
selected_tickers = []
loading = False

def update_selected_display():
    # Clear current display
    for widget in selected_frame.winfo_children():
        widget.destroy()

    global selected_tickers

    max_columns = 5
    for idx, ticker in enumerate(selected_tickers):
        row = idx // max_columns
        col = idx % max_columns

        item_frame = tk.Frame(selected_frame, bg="#333333", padx=5, pady=5)
        item_frame.grid(row=row, column=col, padx=5, pady=5, sticky="w")

        label = tk.Label(item_frame, text=ticker, fg="white", bg="#333333")
        label.pack(side="left")

        remove_btn = tk.Button(item_frame, text="❌", bg="#444444", fg="white", width=2,
                               command=lambda t=ticker: remove_ticker(t))
        remove_btn.pack(side="right", padx=5)

    if len(selected_tickers) != 0:
        selected_frame.place(y=450, x=350, anchor="center")
    else:
        selected_frame.place_forget()


def add_ticker(ticker):
    if ticker not in selected_tickers:
        selected_tickers.append(ticker)
        update_selected_display()

def remove_ticker(ticker):
    if ticker in selected_tickers:
        selected_tickers.remove(ticker)
        update_selected_display()

def message_window(msg, title):
    err_window = ctk.CTkToplevel(window)
    name = "Error!" if title == 1 else "Success!"
    err_window.title(name)
    err_window.geometry("400x200")
    err_window.resizable(False, False)
    err_window.grab_set()  # Modal

    err_message = ctk.CTkLabel(err_window, text=msg, font=FONT_SUBHEADING)
    err_message.pack(pady=(50, 50))

    btn = ctk.CTkButton(err_window, text="OK", command=err_window.destroy)
    btn.pack(pady=(0, 30))

def calc_button_function():
    if not selected_tickers:
        return  # Do nothing if no tickers selected

    # Create a new top-level window
    config_window = ctk.CTkToplevel(window)
    config_window.title("Generate Data")
    config_window.geometry("400x350")
    config_window.resizable(False, False)
    config_window.grab_set()  # Makes this window modal

    # Heading
    heading = ctk.CTkLabel(config_window, text="Select Data Type", font=FONT_SUBHEADING)
    heading.pack(pady=(20, 10))

    # Option Buttons (styled like radio but nicer using CTkSegmentedButton)
    selected_option = ctk.StringVar(value="Beta")  # Default selection

    option_selector = ctk.CTkSegmentedButton(config_window, values=["Beta", "Volatility", "Prices"], variable=selected_option)
    option_selector.pack(pady=10)

    # Date Range Entry Section
    date_frame = ctk.CTkFrame(config_window, fg_color="transparent")
    date_frame.pack(pady=20)

    from_label = ctk.CTkLabel(date_frame, text="From Date (YYYY-MM-DD):", anchor="w", width=180)
    from_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
    from_entry = ctk.CTkEntry(date_frame, width=180, placeholder_text="e.g., 2022-01-01")
    from_entry.grid(row=0, column=1, padx=10, pady=5)

    to_label = ctk.CTkLabel(date_frame, text="To Date (YYYY-MM-DD):", anchor="w", width=180)
    to_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
    to_entry = ctk.CTkEntry(date_frame, width=180, placeholder_text="e.g., 2023-01-01")
    to_entry.grid(row=1, column=1, padx=10, pady=5)

    # Generate Button
    generate_button = ctk.CTkButton(config_window, text="Generate", width=120,
                                     command=lambda: pass_function_placeholder())
    generate_button.pack(pady=20)


    def pass_function_placeholder():
        selected = selected_option.get()
        from_date = from_entry.get()
        to_date = to_entry.get()
        
        def validate_date(date_str, label):
            try:
                # Try parsing date string to datetime
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                return dt
            except ValueError:
                message_window(f"{label} must be in YYYY-MM-DD format and valid.", 1)
                return None
        
        from_dt = validate_date(from_date, "From date")
        if not from_dt:
            return
        
        to_dt = validate_date(to_date, "To date")
        if not to_dt:
            return
        
        # Check if to_date > today
        today = datetime.today()
        if to_dt > today:
            message_window("To date cannot be greater than today's date.", 1)
            return
        
        if not to_dt > from_dt:
            message_window("From date cannot be greater than To Date.", 1)
            return
        # If all valid, call generate
        if selected == "Beta":
            generate_beta(selected_tickers, from_date, to_date)
        elif selected == "Volatility":
            generate_volatility(selected_tickers, from_date, to_date)
        elif selected == "Prices":
            columns = []
            if close_var.get():
                columns.append("Close")
            if volume_var.get():
                columns.append("Volume")
            if not columns:
                message_window("Please select at least one column: Close or Volume.", 1)
                return
            generate_prices(selected_tickers, from_date, to_date, columns)

        message_window("Generated Successfuly!!", 0)
    
    # Frame for checkboxes (only shown if Prices is selected)
    column_frame = ctk.CTkFrame(config_window, fg_color="transparent")
    column_frame.pack(pady=(0, 10))

    # Checkboxes for "Close" and "Volume"
    close_var = ctk.BooleanVar(value=False)
    volume_var = ctk.BooleanVar(value=False)

    close_check = ctk.CTkCheckBox(column_frame, text="Close", variable=close_var)
    volume_check = ctk.CTkCheckBox(column_frame, text="Volume", variable=volume_var)
    
    # Show/hide checkboxes based on selected option
    def update_column_options():
        if selected_option.get() == "Prices":
            close_check.pack(side="left", padx=10)
            volume_check.pack(side="left", padx=10)
        else:
            close_check.pack_forget()
            volume_check.pack_forget()

    selected_option.trace_add("write", lambda *args: update_column_options())
    update_column_options()  # Initial update

    
    # Placeholder for your future logic
        
def threaded_suggestions():
    typed = search_bar.get("0.0", "end").lower().strip()
    query = typed

    if not query:
        suggestion_frame.place_forget()
        return

    # This runs in background
    instance.search(query)
    result_dict = instance.result_dict
    global selected_tickers
    # UI updates must happen in main thread
    def update_ui():
        for wid in suggestion_frame.winfo_children():
            wid.destroy()
        matches = [co for co in result_dict]
        if matches:
            for idx, co in enumerate(matches):
                row_frame = tk.Frame(suggestion_frame, bg=DARK_BG if idx % 2 == 0 else "#262626")
                row_frame.pack(fill="x")

                # Name label (left)
                name_label = tk.Label(row_frame, text=co["name"], bg=row_frame["bg"], fg="white", anchor="w", padx=10)
                name_label.pack(side="left", fill="x", expand=True)
                
                # Small "+" button (right)
                add_button = tk.Button(row_frame, text="+", width=2, height=1, bg="#444", fg="white", relief="flat", command=lambda t=co["ticker"]: add_ticker(t))
                add_button.pack(side="right", padx=10, pady=2)                     
                
                # Ticker label (right)
                ticker_label = tk.Label(row_frame, text=co["ticker"], bg=row_frame["bg"], fg="white", anchor="e", padx=10)
                ticker_label.pack(side="right", fill="x", expand=True)


            # # Position the suggestion box directly below search bar
            suggestion_frame.place(rely=0.38, relx=0.017, width=640, height=200)
        
    window.after(0, update_ui)  # Run UI update in main thread


def update_suggestions():
    thread = threading.Thread(target=threaded_suggestions)
    thread.daemon = True  # Make thread a daemon so it exits when main thread exits
    thread.start()

debounce_id = None  # global variable for debounce

def debounce_update():
    global debounce_id
    if debounce_id:
        search_bar.after_cancel(debounce_id)
    debounce_id = search_bar.after(500, update_suggestions)

     
def suggestion_click():
    mainframe.pack_forget()
    
    # Start the data fetching thread
    
# Main search frame
mainframe = ctk.CTkFrame(window, width=680, height=350, corner_radius=BORDER_RADIUS, fg_color=LIGHT_BG)
mainframe.pack(anchor="n", pady=30, padx=20)
mainframe.pack_propagate(False)

app_title = ctk.CTkLabel(mainframe, text="YFinance Toolkit", font=FONT_HEADING, text_color=TEXT_COLOR)
app_title.place(x=20, y=15)

search_text = ctk.CTkLabel(mainframe, text="Search for a company", font=FONT_SUBHEADING, text_color=SECONDARY_TEXT)
search_text.place(x=20, y=50)


# Create search bar with a modern look
search_frame = ctk.CTkFrame(mainframe, width=640, height=46, corner_radius=BORDER_RADIUS, fg_color=DARK_BG, border_width=1, border_color=ACCENT_COLOR)
search_frame.place(x=10, y=85)

search_icon = ctk.CTkLabel(search_frame, text="🔍", font=("Segoe UI", 20))
search_icon.place(x=10, y=10)

search_bar = ctk.CTkTextbox(search_frame, width=500, height=30, fg_color=DARK_BG, border_width=0, font=FONT_MAIN, text_color=TEXT_COLOR)
search_bar.place(x=40, y=8)
search_bar.bind("<KeyRelease>", lambda event: debounce_update())


# Improved suggestion box with highlighting
suggestion_frame = tk.Frame(mainframe, height=12, width=86, 
                           bg=DARK_BG,
                           bd=0, highlightthickness=1, highlightcolor=ACCENT_COLOR, 
                           relief="flat" )

suggestion_frame.place_forget()

# Dashboard frame with improved styling
dash_board_frame = ctk.CTkFrame(window, width=680, height=700, corner_radius=BORDER_RADIUS, fg_color=DARK_BG)
dash_board_frame.pack_forget()
dash_board_frame.pack_propagate(False)
# Frame to hold selected items
selected_frame = ctk.CTkFrame(window, width=680, height=350, corner_radius=BORDER_RADIUS, fg_color=LIGHT_BG)
selected_frame.place_forget()

calc_button = ctk.CTkButton(window, text="Calculate", width=80, height=30, 
                             command=calc_button_function, fg_color=ACCENT_COLOR, 
                             hover_color=HIGHLIGHT_COLOR, font=FONT_MAIN,
                                )
calc_button.place(relx=0.5, rely=1.0, y=-20, anchor="s")




info_vars = {}

# Company Data

def back_to_search():
    instance.back()
    dash_board_frame.pack_forget()
    search_bar.delete("1.0", tk.END)
    
    mainframe.pack(anchor="n", pady=30)

def on_closing():
    global instance
    instance.driver.quit()
    window.destroy()

window.protocol("WM_DELETE_WINDOW", on_closing)

window.mainloop()