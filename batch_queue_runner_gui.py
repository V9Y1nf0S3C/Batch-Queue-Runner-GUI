# -*- coding: utf-8 -*-
"""
Batch Queue Runner GUI Application (v3 - Dynamic Queue)

This script provides a graphical user interface (GUI) built with Tkinter
for managing and executing a list of scripts (e.g., .bat, .py, .cmd)
concurrently.

Features:
- Add scripts via file dialog or drag-and-drop.
- Specify command-line arguments for scripts (globally or per script).
- Control the maximum number of scripts running in parallel.
- View execution logs in real-time.
- Dynamically add scripts to the queue even while execution is active.
- Visual indication of script status (pending, completed, failed).
- Stop execution gracefully (allows currently running scripts to finish).
- Option to allow or disallow duplicate scripts in the queue.
"""

# --- Standard Library Imports ---
import subprocess
import sys
import os
import queue
import threading
import time
import datetime

# --- GUI Toolkit Imports ---
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext

# --- Third-Party Imports ---
# TkinterDnD2 is required for drag-and-drop functionality.
# Ensure it's installed (`pip install tkinterdnd2`).
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # Provide guidance if the required library is missing.
    print("Error: tkinterdnd2 library not found.")
    print("Please install it using: pip install tkinterdnd2")
    sys.exit(1) # Exit if essential dependency is missing.

# --- Application Constants ---
APP_TITLE = "Batch Queue Runner (v3 - Dynamic Queue)" # Title displayed in the window bar.
DEFAULT_MAX_PARALLEL = 2 # Default value for the maximum number of concurrent scripts.
MONITOR_INTERVAL_MS = 500 # Time in milliseconds for periodic checks (if needed, currently unused).
COMPLETED_COLOR = "gray" # Text color for successfully completed scripts in the listbox.
DEFAULT_COLOR = "black" # Default text color for scripts in the listbox.
FAILED_COLOR = "red" # Text color for failed scripts in the listbox.

# --- Helper Functions ---

def parse_dropped_files(dropped_string: str) -> list[str]:
    """
    Parses the string provided by TkinterDnD during a file drop event.

    Handles cases where filenames might contain spaces and are enclosed
    in curly braces `{}` by TkinterDnD, as well as space-separated paths.

    Args:
        dropped_string: The raw string data from the TkinterDnD event.data.

    Returns:
        A list of potential file or directory paths extracted from the string.
        Returns an empty list if parsing fails or the string is empty.
    """
    paths = [] # Initialize an empty list to store extracted paths.
    # Check if the string contains curly braces, indicating potentially complex paths.
    if '{' in dropped_string and '}' in dropped_string:
        import re # Import regex module only if needed.
        # Find all substrings enclosed in curly braces.
        potential_paths = re.findall(r'\{(.*?)\}', dropped_string)
        # Remove the braced parts and check if there's anything left (space-separated paths).
        remaining_string = re.sub(r'\{.*?\}', '', dropped_string).strip()
        if remaining_string:
            # If there's remaining text, split it by spaces and add to potential paths.
            potential_paths.extend(remaining_string.split())
        # Filter out any empty strings that might result from parsing.
        paths = [p for p in potential_paths if p]
    else:
        # If no braces, assume paths are simply space-separated.
        paths = dropped_string.split()

    # Log the parsed paths for debugging purposes.
    print(f"Parsed dropped paths: {paths}")
    return paths # Return the list of identified paths.

# --- Main Application Class ---

class ScriptExecutorApp:
    """
    Encapsulates the GUI application logic for the Batch Queue Runner.

    Manages the user interface, script queue, execution threads, logging,
    and overall application state.
    """
    def __init__(self, master: TkinterDnD.Tk):
        """
        Initializes the ScriptExecutorApp.

        Args:
            master: The root Tkinter window (specifically a TkinterDnD.Tk instance).
        """
        self.master = master # Store the root window instance.
        self.master.title(APP_TITLE) # Set the window title.
        # Allow horizontal resizing, but disable vertical resizing.
        self.master.resizable(True, False)

        # --- Internal Data Structures ---
        # List to store tuples of (script_path, args_string). Mirrored by listbox.
        self.scripts_in_listbox: list[tuple[str, str]] = []
        # Thread-safe queue to hold tasks (script_path, args_string, listbox_index) for workers.
        self.task_queue: queue.Queue[tuple[str, str, int] | None] = queue.Queue()
        # List to keep references to active worker thread objects.
        self.worker_threads: list[threading.Thread] = []
        # Flag indicating if script execution is currently in progress.
        self.execution_active: bool = False
        # Event object used to signal worker threads to stop processing new tasks.
        self.stop_event: threading.Event = threading.Event()
        # Counter for the number of currently running worker threads.
        self.active_workers_count: int = 0
        # Lock to protect access to the active_workers_count variable.
        self.count_lock: threading.Lock = threading.Lock()
        # Lock to prevent race conditions when checking for final completion.
        self.completion_lock: threading.Lock = threading.Lock()

        # --- GUI Initialization ---
        # Build and arrange all the widgets within the master window.
        self._create_widgets()
        # Configure the script queue listbox to accept dropped files.
        self._setup_drag_drop()
        # Calculate and set the window size and position it in the center of the screen.
        self._center_window()

        # Log the application start event.
        self._log("Application started.")
        # Register a callback function to handle the window close ('X') button event.
        self.master.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _create_widgets(self):
        """
        Creates and arranges all the GUI widgets within the main window.
        Uses ttk widgets for a more modern look and feel where possible.
        """
        global DEFAULT_COLOR # Access the global variable to potentially update it.

        # Create the main container frame with padding.
        main_frame = ttk.Frame(self.master, padding="10")
        # Grid the main frame to expand with the window.
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Configure the master window's column to resize with the window.
        self.master.columnconfigure(0, weight=1)
        # Configure the main frame's column to resize.
        main_frame.columnconfigure(0, weight=1)

        # --- Input Frame (Add Script, Arguments) ---
        input_frame = ttk.Frame(main_frame)
        # Place the input frame at the top.
        input_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))
        # Make the arguments entry field expand horizontally.
        input_frame.columnconfigure(2, weight=1)

        # Button to open the file selection dialog.
        self.add_button = ttk.Button(input_frame, text="Add Script(s)", command=self._add_script_dialog)
        self.add_button.grid(row=0, column=0, padx=(0, 5))

        # Label for the arguments entry field.
        ttk.Label(input_frame, text="Arguments:").grid(row=0, column=1, sticky=tk.W, padx=(5, 5))
        # String variable to hold the current arguments entered by the user.
        self.current_args_var = tk.StringVar()
        # Entry field for users to input arguments.
        self.args_entry = ttk.Entry(input_frame, textvariable=self.current_args_var, width=40)
        self.args_entry.grid(row=0, column=2, sticky=(tk.W, tk.E), padx=(0, 10))

        # Boolean variable linked to the 'Allow Duplicates' checkbox.
        self.allow_duplicates_var = tk.BooleanVar(value=False)
        # Checkbox to control whether duplicate script paths can be added.
        self.duplicates_check = ttk.Checkbutton(input_frame, text="Allow Duplicates", variable=self.allow_duplicates_var)
        self.duplicates_check.grid(row=0, column=3, padx=(5, 0))

        # --- Queue Frame (Listbox for Scripts) ---
        queue_frame = ttk.LabelFrame(main_frame, text="Script Queue (Drag & Drop Files Here)", padding="5")
        # Place the queue frame below the input frame, making it expand.
        queue_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 5))
        # Configure the queue frame's column and row to resize.
        queue_frame.columnconfigure(0, weight=1)
        queue_frame.rowconfigure(0, weight=1) # Allow listbox to expand vertically.

        # The main listbox to display the scripts in the queue. Allows multiple selections.
        self.queue_listbox = tk.Listbox(queue_frame, height=15, width=80, selectmode=tk.EXTENDED)
        self.queue_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Attempt to get the default foreground color from the listbox for consistency.
        try:
            DEFAULT_COLOR = self.queue_listbox.cget("fg")
        except tk.TclError:
             # If getting the color fails (e.g., on some platforms), keep the hardcoded fallback.
             pass # Keep fallback default color defined earlier.

        # Vertical scrollbar for the listbox.
        queue_scrollbar_y = ttk.Scrollbar(queue_frame, orient=tk.VERTICAL, command=self.queue_listbox.yview)
        queue_scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        # Link the listbox's vertical scrolling to the scrollbar.
        self.queue_listbox.config(yscrollcommand=queue_scrollbar_y.set)

        # Horizontal scrollbar for the listbox.
        queue_scrollbar_x = ttk.Scrollbar(queue_frame, orient=tk.HORIZONTAL, command=self.queue_listbox.xview)
        queue_scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        # Link the listbox's horizontal scrolling to the scrollbar.
        self.queue_listbox.config(xscrollcommand=queue_scrollbar_x.set)

        # Frame to hold buttons related to the listbox (Edit Args, Remove).
        listbox_button_frame = ttk.Frame(queue_frame)
        listbox_button_frame.grid(row=2, column=0, columnspan=2, pady=(5,0), sticky=tk.W)

        # Button to edit the arguments of the selected script.
        self.edit_args_button = ttk.Button(listbox_button_frame, text="Edit Args", command=self._edit_selected_args)
        self.edit_args_button.pack(side=tk.LEFT, padx=(0, 5))

        # Button to remove selected scripts from the listbox.
        self.remove_button = ttk.Button(listbox_button_frame, text="Remove Selected", command=self._remove_script)
        self.remove_button.pack(side=tk.LEFT)

        # --- Control Frame (Parallelism, Start/Stop) ---
        control_frame = ttk.LabelFrame(main_frame, text="Execution Control", padding="5")
        # Place the control frame below the queue frame.
        control_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 5))

        # Label for the max parallel spinbox.
        ttk.Label(control_frame, text="Max Parallel Scripts:").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        # Integer variable linked to the spinbox, initialized with the default.
        self.max_parallel_var = tk.IntVar(value=DEFAULT_MAX_PARALLEL)
        # Spinbox to allow the user to set the maximum number of parallel executions.
        self.max_parallel_spinbox = ttk.Spinbox(control_frame, from_=1, to=32, textvariable=self.max_parallel_var, width=5)
        self.max_parallel_spinbox.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))

        # Button to start the execution of scripts in the queue.
        self.start_button = ttk.Button(control_frame, text="Start Execution", command=self._start_execution)
        self.start_button.grid(row=0, column=2, padx=(10, 5))

        # Button to signal the execution to stop (initially disabled).
        self.stop_button = ttk.Button(control_frame, text="Stop Execution", command=self._stop_execution, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=3, padx=(5, 0))

        # --- Log Frame (Scrolled Text Area) ---
        log_frame = ttk.LabelFrame(main_frame, text="Logs", padding="5")
        # Place the log frame below the control frame, allowing it to expand vertically.
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 5))
        # Configure the log frame's column and row to resize.
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        # Configure the main frame's row containing the log to expand vertically.
        main_frame.rowconfigure(3, weight=1) # Make log area vertically resizable.

        # Scrolled text widget to display log messages (initially read-only).
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=80, state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # --- Status Bar ---
        # String variable to hold the current status message.
        self.status_var = tk.StringVar(value="Status: Idle.")
        # Label widget at the bottom acting as a status bar.
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        # Place the status bar at the very bottom, spanning all columns.
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E))

    def _setup_drag_drop(self):
        """Registers the queue listbox as a drop target for files."""
        # Register the listbox to accept file drops (DND_FILES type).
        self.queue_listbox.drop_target_register(DND_FILES)
        # Bind the drop event ('<<Drop>>') to the _handle_drop method.
        self.queue_listbox.dnd_bind('<<Drop>>', self._handle_drop)

    def _center_window(self):
        """Calculates window dimensions and positions it in the center of the screen."""
        # Ensure Tkinter has processed pending geometry changes.
        self.master.update_idletasks()
        # Get the minimum required width and height based on widget content.
        min_width = self.master.winfo_reqwidth()
        min_height = self.master.winfo_reqheight()
        # Set the minimum size of the window to prevent it from becoming too small.
        self.master.minsize(min_width, min_height)
        # Get the screen width and height.
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        # Calculate a reasonable initial width (half screen width, but not less than min).
        width = max(min_width, screen_width // 2)
        # Calculate a reasonable initial height (half screen height, but not less than min).
        height = max(min_height, screen_height // 2)
        # Calculate the x-coordinate to center the window horizontally.
        x = (screen_width // 2) - (width // 2)
        # Calculate the y-coordinate to center the window vertically.
        y = (screen_height // 2) - (height // 2)
        # Set the window's geometry (size and position).
        self.master.geometry(f'{width}x{height}+{x}+{y}')

    def _insert_log_message(self, message: str):
        """
        Safely inserts a message into the log ScrolledText widget.

        This method is designed to be called from the main GUI thread,
        often scheduled via `master.after`. It handles potential errors
        if the widget is destroyed before the update occurs.

        Args:
            message: The string message to insert into the log.
        """
        try:
            # Temporarily enable the text widget to allow insertion.
            self.log_text.config(state=tk.NORMAL)
            # Insert the message at the end, followed by a newline.
            self.log_text.insert(tk.END, message + "\n")
            # Automatically scroll the text widget to show the latest message.
            self.log_text.see(tk.END)
            # Disable the text widget again to make it read-only.
            self.log_text.config(state=tk.DISABLED)
        except tk.TclError:
            # Handle the case where the widget might have been destroyed.
            print("Log Error: Could not write to log widget (already destroyed?)")
        except Exception as e:
             # Catch any other unexpected errors during log insertion.
             print(f"Unexpected error inserting log message: {e}")


    def _log(self, message: str):
        """
        Logs a message to both the console and the GUI log widget.

        Prepends a timestamp to the message. Ensures GUI updates happen
        on the main thread using `master.after`.

        Args:
            message: The message string to log.
        """
        # Get the current time and format it.
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Format the final log entry with the timestamp.
        log_entry = f"[{now}] {message}"
        # Print the log entry to the console (useful for debugging).
        print(log_entry)
        try:
            # Check if the main window still exists before scheduling the update.
            if self.master.winfo_exists():
                # Schedule the _insert_log_message method to run in the main GUI thread.
                # `after(0, ...)` makes it run as soon as possible in the event loop.
                self.master.after(0, self._insert_log_message, log_entry)
        except Exception as e:
            # Log any errors that occur during the scheduling process.
            print(f"Error scheduling log update: {e}")


    def _update_status(self, message: str):
        """
        Updates the text in the status bar label.

        Ensures the update happens safely in the main GUI thread using `master.after`.

        Args:
            message: The new message to display in the status bar.
        """
        def update():
            """Inner function to perform the actual update, executed by `after`."""
            try:
                # Check if the master window still exists.
                if self.master.winfo_exists():
                    # Set the text of the status bar variable.
                    self.status_var.set(f"Status: {message}")
            except tk.TclError:
                # Handle error if the widget/variable is destroyed before update.
                print("Status Update Error: Could not set status var (window destroyed?)")
            except Exception as e:
                # Catch any other unexpected errors during status update.
                print(f"Unexpected error updating status: {e}")

        try:
            # Check if the master window exists before scheduling.
            if self.master.winfo_exists():
                # Schedule the inner `update` function to run in the main GUI thread.
                self.master.after(0, update)
        except Exception as e:
            # Log errors during the scheduling process itself.
            print(f"Error scheduling status update: {e}")

    def _add_script_dialog(self):
        """
        Opens a standard file dialog to allow the user to select script files.

        Adds the selected files to the queue using the currently entered arguments.
        """
        # Open the 'askopenfilenames' dialog, allowing multiple file selections.
        file_paths = filedialog.askopenfilenames(
            title="Select Scripts (.bat, .py, ...)",
            # Define file type filters for common script/executable types.
            filetypes=[("Script/Batch Files", "*.bat *.cmd *.ps1 *.py *.sh"),
                       ("Executable Files", "*.exe"),
                       ("All Files", "*.*")]
        )
        # Proceed only if the user selected one or more files.
        if file_paths:
            # Get the arguments currently entered in the arguments entry field.
            args = self.current_args_var.get()
            # Call the common method to add the selected scripts to the list/queue.
            self._add_scripts_to_list(file_paths, args)

    def _handle_drop(self, event):
        """
        Callback function executed when files are dropped onto the listbox.

        Parses the dropped data to extract file paths and adds them to the queue.

        Args:
            event: The event object containing data about the drop (event.data).
        """
        # Log the raw drop event data for debugging.
        self._log(f"Drop event detected. Data: '{event.data}'")
        # Use the helper function to parse file paths from the event data string.
        file_paths = parse_dropped_files(event.data)
        # Proceed only if valid file paths were extracted.
        if file_paths:
            # Get the arguments currently entered in the arguments entry field.
            args = self.current_args_var.get()
            # Call the common method to add the dropped scripts to the list/queue.
            self._add_scripts_to_list(file_paths, args)
        else:
            # If parsing failed, log an error and show a warning to the user.
            error_msg = f"Could not parse valid file paths from drop: '{event.data}'"
            self._log(f"Error: {error_msg}")
            messagebox.showwarning("Drop Error", error_msg)

    def _add_scripts_to_list(self, file_paths: list[str], args_string: str):
        """
        Adds a list of script file paths to the internal list, the GUI listbox,
        and potentially the active execution queue if running.

        Handles duplicate checking and path validation.

        Args:
            file_paths: A list of strings, where each string is a path to a script.
            args_string: The command-line arguments string to associate with these scripts.
        """
        added_count = 0 # Counter for successfully added scripts.
        # Get the current state of the 'Allow Duplicates' checkbox.
        allow_duplicates = self.allow_duplicates_var.get()
        # Create a set of existing absolute paths for efficient duplicate checking (if needed).
        # This uses the first element (path) from each tuple in self.scripts_in_listbox.
        existing_paths = {item[0] for item in self.scripts_in_listbox}

        # Iterate through each file path provided.
        for file_path in file_paths:
            # Convert the path to an absolute path for consistency.
            abs_path = os.path.abspath(file_path)
            # Check if the file actually exists.
            if not os.path.exists(abs_path):
                # Log a warning and skip this file if it doesn't exist.
                self._log(f"Warning: Skipped non-existent file: {abs_path}")
                continue # Move to the next file path.

            # Check for duplicates if 'Allow Duplicates' is not checked.
            if not allow_duplicates and abs_path in existing_paths:
                # Log that a duplicate was skipped.
                self._log(f"Skipped duplicate: {abs_path}")
                continue # Move to the next file path.

            # --- Add the script ---
            # 1. Add to the internal data structure: Store as (absolute_path, arguments).
            self.scripts_in_listbox.append((abs_path, args_string))
            # 2. Get the index for the listbox (which is the new size - 1).
            #    This index is crucial for linking the listbox item to the task later.
            listbox_index = len(self.scripts_in_listbox) - 1

            # 3. Format the text to be displayed in the listbox.
            display_text = f"{abs_path}" # Start with the absolute path.
            if args_string:
                # Append arguments in brackets if they exist.
                display_text += f"  [{args_string}]"

            # 4. Add the formatted text to the GUI listbox.
            self.queue_listbox.insert(tk.END, display_text)
            # 5. Set the initial text color for the new item.
            self.queue_listbox.itemconfig(tk.END, {'fg': DEFAULT_COLOR})

            # 6. Add the path to the set for duplicate checking in this loop iteration.
            existing_paths.add(abs_path)
            # 7. Increment the counter for added scripts.
            added_count += 1

            # 8. Add to active queue if execution is running (Dynamic Queue Update).
            if self.execution_active:
                # Create the task tuple: (path, args, original_listbox_index).
                task = (abs_path, args_string, listbox_index)
                # Put the task onto the thread-safe queue for a worker to pick up.
                self.task_queue.put(task)
                # Log that the task was added dynamically.
                self._log(f"Added task to active queue (Index {listbox_index}): {os.path.basename(abs_path)}")
                # Optional: Could update status here, but might be too frequent.

        # After processing all paths, update logs and status based on how many were added.
        if added_count > 0:
            log_msg = f"Added {added_count} script(s) to list."
            # Append info about adding to the active queue if relevant.
            if self.execution_active:
                 log_msg += " (Also added to active task queue)"
            self._log(log_msg)
            self._update_status(f"Added {added_count} script(s).")
            # Scroll the listbox to show the newly added items.
            self.queue_listbox.see(tk.END)
        else:
            # Log if no new scripts were added (e.g., all were duplicates or invalid).
            log_msg = "No new scripts added (duplicates skipped or files invalid)."
            self._log(log_msg)
            # Avoid overwriting a potentially more informative status message (like "Running...").


    def _remove_script(self):
        """
        Removes the currently selected script(s) from the listbox and internal list.

        Warns the user if attempting removal while execution is active, as this
        does not stop already running or queued tasks.
        """
        # Get the indices of the items currently selected in the listbox.
        selected_indices = self.queue_listbox.curselection()
        # If nothing is selected, show a warning and return.
        if not selected_indices:
            messagebox.showwarning("Warning", "No scripts selected to remove.")
            return

        removed_paths_basenames = [] # List to store basenames of removed scripts for logging.

        # --- Warn if removing during active execution ---
        if self.execution_active:
            # Ask for confirmation because removal is only visual/preventative for future runs.
            if not messagebox.askyesno("Confirm Removal During Execution",
                                       "Execution is active.\n"
                                       "Removing items from the list will NOT stop them if they are "
                                       "already running or queued.\n"
                                       "It only removes them visually and prevents future runs in THIS session.\n\n"
                                       "Continue removal?"):
                # If the user clicks 'No', abort the removal process.
                return

        # Iterate through the selected indices in reverse order.
        # This is important to avoid index shifting issues when deleting multiple items.
        for i in reversed(selected_indices):
            try:
                # 1. Remove the script tuple from the internal list using the index.
                removed_path, _ = self.scripts_in_listbox.pop(i)
                # 2. Delete the corresponding item from the GUI listbox.
                self.queue_listbox.delete(i)
                # 3. Store the basename for the log message.
                removed_paths_basenames.append(os.path.basename(removed_path))
                # NOTE: We intentionally do *not* try to remove items from the
                # `self.task_queue` here. It's complex to do safely and reliably
                # while threads are actively consuming from it (race conditions).
                # The warning to the user covers this behavior.
            except IndexError:
                # Log an error if the index is somehow invalid during removal.
                self._log(f"Error: Index {i} out of bounds during removal.")
            except Exception as e:
                # Log any other unexpected errors during removal.
                self._log(f"Unexpected error removing item at index {i}: {e}")


        # If any scripts were successfully removed, log and update the status.
        if removed_paths_basenames:
            log_msg = f"Removed {len(removed_paths_basenames)} script(s) from list: {', '.join(removed_paths_basenames)}"
            self._log(log_msg)
            self._update_status(f"Removed {len(removed_paths_basenames)} script(s) from list.")

    def _edit_selected_args(self):
        """
        Opens a dialog to edit the arguments associated with the single selected script.

        Warns the user if editing during execution, as changes only affect the
        list representation and future potential runs, not currently active tasks.
        """
        # Get the indices of selected items in the listbox.
        selected_indices = self.queue_listbox.curselection()
        # Check if exactly one item is selected.
        if not selected_indices:
            messagebox.showwarning("Warning", "No script selected to edit arguments for.")
            return
        if len(selected_indices) > 1:
            messagebox.showwarning("Warning", "Please select only one script to edit its arguments.")
            return

        # Get the index of the single selected item.
        index = selected_indices[0]

        # --- Warn if editing during active execution ---
        if self.execution_active:
            messagebox.showinfo("Edit Args During Execution",
                                "Execution is active.\n"
                                "Editing arguments now will only affect the script's representation in the list "
                                "and any FUTURE runs (e.g., if execution is restarted or the script added again).\n"
                                "It will NOT change arguments for tasks already queued or currently running.")

        try:
            # Get the current path and arguments from the internal list using the index.
            current_path, current_args = self.scripts_in_listbox[index]
            # Open a simple dialog box asking the user for the new arguments.
            # Prefill the dialog with the current arguments.
            new_args = simpledialog.askstring(
                "Edit Arguments",
                f"Enter new arguments for:\n{os.path.basename(current_path)}",
                initialvalue=current_args
            )

            # Proceed only if the user entered new arguments (didn't cancel).
            # new_args will be None if the user cancels. An empty string is valid.
            if new_args is not None:
                # 1. Update the arguments in the internal list.
                self.scripts_in_listbox[index] = (current_path, new_args)

                # 2. Update the display text in the GUI listbox.
                # Reconstruct the display text with the new arguments.
                display_text = f"{current_path}" + (f"  [{new_args}]" if new_args else "")
                # Remember the current text color (e.g., if it was completed/failed).
                original_color = self.queue_listbox.itemcget(index, 'fg')
                # Delete the old listbox item.
                self.queue_listbox.delete(index)
                # Insert the updated item back at the same index.
                self.queue_listbox.insert(index, display_text)
                # Restore the original text color.
                self.queue_listbox.itemconfig(index, {'fg': original_color})
                # Re-select the edited item for user convenience.
                self.queue_listbox.selection_set(index)

                # 3. Log the change and update the status bar.
                log_msg = (f"Updated arguments for list item {index} "
                           f"({os.path.basename(current_path)}) to: [{new_args}]")
                self._log(log_msg)
                self._update_status(f"Updated arguments for selected script in list.")

        except IndexError:
            # Handle error if the selected index is somehow invalid.
            err_msg = "Error: Selected script index not found in internal list during edit."
            self._log(err_msg)
            messagebox.showerror("Error", err_msg)
        except Exception as e:
            # Handle any other unexpected errors during the edit process.
            err_msg = f"Unexpected error editing arguments for index {index}: {e}"
            self._log(err_msg)
            messagebox.showerror("Error", err_msg)


    def _start_execution(self):
        """
        Initiates the script execution process.

        - Checks prerequisites (execution not already active, queue not empty).
        - Validates the 'Max Parallel' setting.
        - Clears the task queue and repopulates it from the listbox.
        - Resets listbox item colors (except for previously failed/completed).
        - Updates GUI element states (disables Start, enables Stop, etc.).
        - Starts the worker threads.
        """
        # Prevent starting if execution is already underway.
        if self.execution_active:
            self._log("Info: Start command ignored, execution already active.")
            messagebox.showinfo("Info", "Execution is already active.")
            return

        # Prevent starting if there are no scripts in the list.
        if not self.scripts_in_listbox:
            self._log("Warning: Start command ignored, script queue is empty.")
            messagebox.showwarning("Warning", "Script queue is empty. Add scripts first.")
            return

        # Get and validate the maximum number of parallel scripts.
        try:
            max_parallel = self.max_parallel_var.get()
            if max_parallel <= 0:
                self._log("Error: Max parallel scripts must be greater than 0.")
                messagebox.showerror("Error", "Max parallel scripts must be greater than 0.")
                return
        except tk.TclError:
             # Handle potential error if the spinbox value is invalid.
             self._log("Error: Invalid value for Max parallel scripts.")
             messagebox.showerror("Error", "Invalid value for Max parallel scripts. Please enter a number.")
             return


        # --- Begin Execution Setup ---
        self.execution_active = True # Set the flag indicating execution has started.
        self.stop_event.clear() # Ensure the stop signal is not set from a previous run.
        self.worker_threads = [] # Clear any old worker thread references.
        self.active_workers_count = 0 # Reset the active worker counter.

        # Clear any leftover tasks from a previous run (or dynamically added tasks not yet run).
        # This ensures we start fresh based on the current listbox content.
        while not self.task_queue.empty():
            try:
                self.task_queue.get_nowait() # Remove an item without blocking.
                self.task_queue.task_done() # Mark it as done (needed for queue joining if used).
            except queue.Empty:
                break # Stop if the queue becomes empty.
            except Exception as e:
                self._log(f"Minor error clearing task queue: {e}") # Log unexpected queue errors.
                break

        # --- Populate Task Queue and Reset Listbox Colors ---
        # Iterate through all scripts currently in the internal list.
        for i, (script_path, args_string) in enumerate(self.scripts_in_listbox):
            try:
                # Get the current color of the listbox item.
                current_color = self.queue_listbox.itemcget(i, 'fg')
                # Reset the color to default *unless* it was already marked as failed or completed
                # in a previous (potentially partial) run. This preserves their status visually.
                if current_color not in (FAILED_COLOR, COMPLETED_COLOR):
                    self.queue_listbox.itemconfig(i, {'fg': DEFAULT_COLOR})
                # Add the task (path, args, original_index) to the queue for workers.
                self.task_queue.put((script_path, args_string, i))
            except tk.TclError:
                 # Handle cases where listbox item might not exist (shouldn't happen here ideally)
                 self._log(f"Warning: Could not access or update listbox item at index {i} during start.")
            except Exception as e:
                 self._log(f"Error processing item {i} ('{os.path.basename(script_path)}') during start: {e}")


        # Get the total number of tasks added to the queue.
        total_tasks = self.task_queue.qsize()
        self._log(f"Starting execution: {total_tasks} initial tasks, Max Parallel: {max_parallel}")
        self._update_status(f"Starting execution (Tasks: {total_tasks}, Max: {max_parallel})...")

        # --- Update GUI State for Active Execution ---
        self.start_button.config(state=tk.DISABLED) # Disable Start button.
        self.stop_button.config(state=tk.NORMAL)   # Enable Stop button.
        # Keep Add/Edit/Remove controls enabled to allow dynamic queue modification.
        self.add_button.config(state=tk.NORMAL)
        self.duplicates_check.config(state=tk.NORMAL)
        self.args_entry.config(state=tk.NORMAL)
        self.remove_button.config(state=tk.NORMAL)
        self.edit_args_button.config(state=tk.NORMAL)
        # Disable changing max parallelism while running to avoid complexity.
        self.max_parallel_spinbox.config(state=tk.DISABLED)

        # --- Launch Worker Threads ---
        # Determine the number of worker threads to start.
        # It's the minimum of the user's setting and the actual number of tasks.
        num_workers_to_start = min(max_parallel, total_tasks)
        # Safety check: Ensure at least one worker if there are tasks (shouldn't be needed with current logic).
        if num_workers_to_start == 0 and total_tasks > 0:
             num_workers_to_start = 1

        # Initialize the active worker count (protected by lock, but okay here before threads start).
        self.active_workers_count = num_workers_to_start

        self._log(f"Launching {num_workers_to_start} worker threads.")
        # Create and start the worker threads.
        for i in range(num_workers_to_start):
            # Create a thread targeting the _worker_thread method.
            # Pass the worker's ID for logging purposes.
            # Set daemon=True so threads exit automatically if the main program exits abruptly.
            thread = threading.Thread(target=self._worker_thread, args=(i,), daemon=True)
            # Store a reference to the thread object.
            self.worker_threads.append(thread)
            # Start the thread's execution.
            thread.start()

        # Handle the edge case where the list was populated but somehow resulted in 0 tasks
        # or 0 workers allowed (e.g., max_parallel was invalid initially but bypassed check - unlikely).
        if total_tasks == 0:
             self._log("No tasks to execute after initialization.")
             # Schedule the finish routine shortly after, as there's nothing to run.
             self.master.after(100, self._check_final_completion)


    def _stop_execution(self):
        """
        Signals the worker threads to stop processing new tasks from the queue.

        Sets the `stop_event` and puts `None` sentinels into the queue
        to potentially wake up blocking `queue.get()` calls. Disables the
        Stop button to prevent multiple clicks.
        """
        # Do nothing if execution isn't currently active.
        if not self.execution_active:
            self._log("Stop command ignored: Execution not active.")
            return

        self._log("Stop signal sent. Waiting for currently active scripts to finish...")
        self._update_status("Stop signal sent. Finishing active scripts...")
        # Set the event that worker threads check periodically.
        self.stop_event.set()
        # Disable the stop button to prevent sending the signal multiple times.
        self.stop_button.config(state=tk.DISABLED)

        # Put sentinel values (None) into the queue.
        # This helps wake up any worker threads that might be blocked waiting
        # indefinitely on `task_queue.get()`. Put slightly more than the number
        # of workers to be safe.
        for _ in range(len(self.worker_threads) + 1):
            try:
                self.task_queue.put(None)
            except Exception as e:
                # Log minor errors during sentinel placement (e.g., queue full, though unlikely).
                self._log(f"Minor error putting sentinel in queue during stop: {e}")

    def _worker_thread(self, worker_id: int):
        """
        The main function executed by each worker thread.

        Continuously fetches tasks from the `task_queue`, executes the
        corresponding script using `subprocess.Popen`, and updates the GUI
        (via `master.after`) upon completion or failure. Stops when the
        `stop_event` is set or the queue provides a `None` sentinel.

        Args:
            worker_id: A unique identifier for this worker thread (used for logging).
        """
        self._log(f"Worker {worker_id}: Started.")

        # Loop indefinitely until explicitly broken.
        while True:
            # Check if the stop signal has been set before attempting to get a task.
            if self.stop_event.is_set():
                self._log(f"Worker {worker_id}: Stop event detected at loop start. Exiting.")
                break

            task = None # Initialize task variable for the current iteration.
            try:
                # Attempt to get a task from the queue.
                # Use a timeout (e.g., 0.5 seconds) so the loop doesn't block indefinitely.
                # This allows the `stop_event` check at the start of the loop to be effective
                # even if the queue remains empty for a while.
                task = self.task_queue.get(timeout=0.5)

                # --- Check for Sentinel ---
                if task is None:
                    # Received the sentinel value (None), indicating a stop request or queue exhaustion signal.
                    self._log(f"Worker {worker_id}: Received stop sentinel (None task). Exiting loop.")
                    # Ensure task_done is called even for the sentinel if it was put deliberately.
                    # self.task_queue.task_done() # Should be called outside loop in finally ideally
                    break # Exit the main `while` loop.

                # --- Process Valid Task ---
                # Unpack the task tuple.
                script_path, args_string, listbox_index = task
                base_name = os.path.basename(script_path) # Get filename for logging.
                self._log(f"Worker {worker_id}: Starting script (Index {listbox_index}): '{base_name}' Args: [{args_string}]")

                process = None # Initialize process variable.
                exit_code = -1   # Default exit code if execution fails early.
                try:
                    # --- Execute Script ---
                    # Quote the script path to handle spaces correctly in the command line.
                    quoted_path = f'"{script_path}"'
                    # Construct the command string using `start /wait`.
                    # `start /wait` ensures the `cmd` window waits for the script to finish
                    # before `Popen` returns. The title helps identify the window if needed.
                    # NOTE: This relies on the Windows `start` command. Might need adjustment for cross-platform use.
                    # NOTE: We redirect stdout/stderr to DEVNULL to prevent script output
                    #       from flooding the console where this GUI app was launched.
                    #       Consider capturing output if needed for detailed logging later.
                    cmd_string = f'start "Worker {worker_id}: {base_name}" /WAIT {quoted_path} {args_string}'
                    self._log(f"Worker {worker_id}: Executing command: {cmd_string}")

                    # Launch the script in a new process. `shell=True` is needed for `start`.
                    # Use DEVNULL to hide the script's own console output/errors from *this* application's console.
                    process = subprocess.Popen(cmd_string, shell=True,
                                               stdout=subprocess.DEVNULL,
                                               stderr=subprocess.DEVNULL,
                                               # `creationflags=subprocess.CREATE_NO_WINDOW` could potentially hide
                                               # the intermediate cmd window opened by 'start', but can sometimes
                                               # interfere with script execution depending on the script. Test carefully.
                                               # creationflags=subprocess.CREATE_NO_WINDOW
                                               )
                    # Wait for the launched process to complete.
                    process.wait()
                    # Get the exit code returned by the script/process.
                    exit_code = process.returncode
                    self._log(f"Worker {worker_id}: Script '{base_name}' (Index {listbox_index}) finished. Exit Code: {exit_code}")

                    # --- Update GUI (Success) ---
                    # Check the stop event *again* after the script finishes. If stop was called
                    # *during* script execution, we don't want to mark it as normally completed.
                    if not self.stop_event.is_set():
                        # Schedule the _mark_completed function to run in the main GUI thread.
                        self.master.after(0, self._mark_completed, listbox_index, exit_code, base_name)

                except FileNotFoundError:
                    # Handle error if the 'start' command itself is not found (unlikely on Windows).
                    error_msg = "Critical Error: 'start' command not found. This tool requires a Windows environment."
                    self._log(error_msg)
                    # Show error message in the GUI (scheduled).
                    self.master.after(0, messagebox.showerror, "Execution Error", error_msg)
                    # Signal all other threads to stop as well, as execution is fundamentally broken.
                    self.stop_event.set()
                    # Mark this specific script as failed in the GUI (scheduled).
                    self.master.after(0, self._mark_failed, listbox_index, base_name, "CmdNotFound")
                    break # Exit worker loop after critical error.

                except Exception as e:
                    # Catch any other exceptions during Popen, wait, or processing.
                    error_msg = f"Error executing '{base_name}' (Index {listbox_index}): {e}"
                    self._log(error_msg)
                    # Show a generic error message in the GUI (scheduled).
                    # Avoid showing overly technical details directly to the user unless necessary.
                    self.master.after(0, messagebox.showerror, "Execution Error", f"Error occurred while running {base_name}:\n{type(e).__name__}")
                    # Mark this script as failed in the GUI (scheduled).
                    self.master.after(0, self._mark_failed, listbox_index, base_name, f"ExecError: {type(e).__name__}")
                    # Note: We typically don't stop all threads for a single script error,
                    # allowing other scripts to continue. Set stop_event here if that's desired.
                    # self.stop_event.set()

            except queue.Empty:
                # --- Handle Empty Queue ---
                # The queue was empty during the `get(timeout=0.5)` call.
                # Check if execution should genuinely stop or if we should just wait longer.
                if self.stop_event.is_set() or not self.execution_active:
                    # If stop is signaled OR the main app logic has marked execution as inactive, exit the loop.
                    self._log(f"Worker {worker_id}: Queue empty and stop signal set or execution inactive. Exiting loop.")
                    break
                else:
                    # Otherwise, execution is still active and no stop signal. The queue might
                    # get more items dynamically. Continue the loop to wait again.
                    # self._log(f"Worker {worker_id}: Queue empty, but execution active. Waiting for more tasks...")
                    continue # Go back to the start of the while loop.

            except Exception as e:
                # --- Handle Critical Worker Errors ---
                # Catch unexpected errors in the main worker loop logic itself.
                self._log(f"CRITICAL ERROR in worker {worker_id} main loop: {e}")
                # Optionally mark any task currently being processed (if `task` is not None) as failed.
                if task:
                     try:
                         script_path, args_string, listbox_index = task
                         base_name = os.path.basename(script_path)
                         self.master.after(0, self._mark_failed, listbox_index, base_name, f"WorkerLoopError: {type(e).__name__}")
                     except Exception as inner_e:
                          self._log(f"Error trying to mark task failed after worker loop error: {inner_e}")
                break # Exit the worker loop due to the critical error.

            finally:
                 # --- Task Completion Signal ---
                 # Crucial: Signal to the queue that the task (if one was retrieved) is done.
                 # This is necessary for `queue.join()` if it were used, but also good practice.
                 # Ensure this runs even if exceptions occurred during task processing.
                 if task is not None: # Only call task_done if a task was actually retrieved
                      try:
                          self.task_queue.task_done()
                      except ValueError:
                          # Can happen if task_done() is called too many times. Log and ignore.
                          self._log(f"Worker {worker_id}: ValueError on task_done (task may have already been marked done).")
                      except Exception as e:
                          self._log(f"Worker {worker_id}: Unexpected error calling task_done: {e}")


        # --- Worker Thread Cleanup ---
        self._log(f"Worker {worker_id}: Finishing.")
        # Use a lock to safely decrement the global counter for active workers.
        with self.count_lock:
            self.active_workers_count -= 1
            # Read the count *after* decrementing.
            active_count = self.active_workers_count
            self._log(f"Worker {worker_id}: Decremented active count. Remaining active workers: {active_count}")

            # --- Check if This Might Be the Last Worker ---
            # If the active count has reached zero or less (shouldn't be less, but check <= 0),
            # this *might* be the last worker finishing its current workload.
            if active_count <= 0:
                # Attempt to acquire the completion lock *without blocking*.
                # This acts as a flag to ensure only one thread triggers the final check
                # when the count first hits zero.
                if self.completion_lock.acquire(blocking=False):
                    self._log(f"Worker {worker_id}: Acquired completion lock as active count reached zero.")
                    try:
                        # Schedule the final completion check function (`_check_final_completion`)
                        # to run in the main GUI thread shortly. This function will verify
                        # if the queue is *also* empty and if execution should truly end.
                        # Using `after(100)` gives a small buffer for other threads/GUI updates.
                        self.master.after(100, self._check_final_completion)
                    except Exception as e:
                        self._log(f"Error scheduling final completion check from worker {worker_id}: {e}")
                        # If scheduling fails, we should release the lock so another attempt can be made.
                        try:
                            self.completion_lock.release()
                            self._log(f"Worker {worker_id}: Released completion lock due to scheduling error.")
                        except (threading.ThreadError, RuntimeError): pass # Ignore if already released
                else:
                    # Another thread already acquired the lock and presumably scheduled the check.
                    self._log(f"Worker {worker_id}: Completion lock already held; another thread will check completion.")
            # If active_count > 0, do nothing special here; other workers are still running.


    def _check_final_completion(self):
        """
        Checks if the execution cycle should be considered fully completed.

        This function runs in the main GUI thread (scheduled by the last worker).
        It verifies that:
        1. The active worker count is zero.
        2. The task queue is empty.
        OR
        3. The stop event has been signaled.

        If conditions are met and execution is marked as active, it calls
        `_on_all_workers_finished` to perform cleanup. It manages the
        `completion_lock` to ensure cleanup happens only once per cycle.
        """
        self._log("Running final completion check (scheduled by last worker)...")

        # Safely read the current active worker count.
        with self.count_lock:
            active_workers = self.active_workers_count

        # Check if the task queue is currently empty.
        queue_empty = self.task_queue.empty()
        # Determine if conditions for finishing are met.
        stop_signaled = self.stop_event.is_set()
        # Finish if no workers AND queue empty, OR if stop was called.
        should_finish = (active_workers <= 0 and queue_empty) or stop_signaled

        # Only proceed with finishing logic if execution is still marked as active.
        if should_finish and self.execution_active:
            self._log(f"Completion Check: Conditions met (Workers: {active_workers}, Queue Empty: {queue_empty}, Stop Set: {stop_signaled}). Finalizing execution.")
            # Call the main cleanup and state reset function.
            # The completion lock acquired by the worker is implicitly released
            # inside _reset_gui_state (which is called by _on_all_workers_finished).
            self._on_all_workers_finished()
        elif self.execution_active:
            # Conditions not met, but execution is technically still active (e.g., workers finished
            # but items were dynamically added to the queue *after* the last worker checked it).
            self._log(f"Completion Check: Conditions NOT met, but execution still marked active (Workers: {active_workers}, Queue Empty: {queue_empty}, Stop Set: {stop_signaled}). Execution cycle continues.")
            # Release the completion lock so that if another worker finishes later
            # (after processing newly added tasks), it can trigger this check again.
            try:
                self.completion_lock.release()
                self._log("Completion Check: Released completion lock as conditions not met.")
            except (threading.ThreadError, RuntimeError): pass # Ignore if not held or already released.
        else:
            # Execution was already marked as inactive before this check ran.
            self._log("Completion Check: Execution already marked inactive. Ensuring lock is released.")
            # Ensure the lock is released just in case.
            try:
                self.completion_lock.release()
            except (threading.ThreadError, RuntimeError): pass # Ignore if not held or already released.


    def _mark_completed(self, listbox_index: int, exit_code: int, base_name: str):
        """
        Updates a listbox item's appearance to indicate successful completion.

        Appends the exit code to the text and changes the color. Runs in the
        main GUI thread via `master.after`.

        Args:
            listbox_index: The index of the item in the listbox to update.
            exit_code: The exit code returned by the script process.
            base_name: The base name of the script file (for logging/status).
        """
        try:
            # Check if the main window still exists before touching widgets.
            if not self.master.winfo_exists(): return

            # Define the suffix to add to the listbox text.
            status_suffix = f" (Done, Code: {exit_code})"
            # Get the current text of the listbox item.
            current_text = self.queue_listbox.get(listbox_index)
            # Remove any previous status suffixes (Done or Failed) to prevent duplication.
            base_text = current_text.split(" (Done")[0].split(" (Failed")[0]
            # Create the new text string.
            new_text = base_text + status_suffix

            # Update the listbox item: delete old, insert new at same index.
            self.queue_listbox.delete(listbox_index)
            self.queue_listbox.insert(listbox_index, new_text)
            # Change the text color to the 'completed' color.
            self.queue_listbox.itemconfig(listbox_index, {'fg': COMPLETED_COLOR})

            # --- Update Status Bar ---
            # Get current queue size and active worker count (read count safely).
            remaining = self.task_queue.qsize()
            with self.count_lock: # Ensure reading count is safe
                 active = max(0, self.active_workers_count) # Use max(0,...) defensively
            max_allowed = self.max_parallel_var.get()
            self._update_status(f"Running: {active}/{max_allowed}, Queue: {remaining}, Finished: '{base_name}' (Code {exit_code})")

        except (tk.TclError, IndexError) as e:
            # Handle errors if the listbox item doesn't exist at the given index
            # (e.g., it was removed by the user after the task started).
            self._log(f"Info: Could not update listbox item at index {listbox_index} for completed script '{base_name}'. Item might have been removed. Error: {e}")
        except Exception as e:
            # Catch any other unexpected errors during the update.
            self._log(f"Error marking completed for index {listbox_index} ('{base_name}'): {e}")


    def _mark_failed(self, listbox_index: int, base_name: str, reason: str = "Error"):
        """
        Updates a listbox item's appearance to indicate failure.

        Appends a failure reason to the text and changes the color to red.
        Runs in the main GUI thread via `master.after`.

        Args:
            listbox_index: The index of the item in the listbox to update.
            base_name: The base name of the script file (for logging/status).
            reason: A short string indicating the reason for failure.
        """
        try:
             # Check if the main window still exists.
            if not self.master.winfo_exists(): return

            # Define the suffix indicating failure.
            status_suffix = f" (Failed: {reason})"
            # Get current text and strip any previous status suffixes.
            current_text = self.queue_listbox.get(listbox_index)
            base_text = current_text.split(" (Done")[0].split(" (Failed")[0]
            # Create the new text string.
            new_text = base_text + status_suffix

            # Update the listbox item.
            self.queue_listbox.delete(listbox_index)
            self.queue_listbox.insert(listbox_index, new_text)
            # Change the text color to the 'failed' color.
            self.queue_listbox.itemconfig(listbox_index, {'fg': FAILED_COLOR})

            # Update the status bar to indicate the failure.
            self._update_status(f"Failed: '{base_name}' (Index {listbox_index}) Reason: {reason}")

        except (tk.TclError, IndexError) as e:
            # Handle cases where the item might have been removed.
            self._log(f"Info: Could not update listbox item at index {listbox_index} for failed script '{base_name}'. Item might have been removed. Error: {e}")
        except Exception as e:
            # Catch any other unexpected errors.
            self._log(f"Error marking failed for index {listbox_index} ('{base_name}'): {e}")


    def _on_all_workers_finished(self):
        """
        Performs final actions when an execution cycle completes.

        This is called by `_check_final_completion` when conditions are met
        (workers finished AND queue empty, OR stop signaled). It marks execution
        as inactive, updates the status bar, resets GUI controls, and shows
        a final confirmation message.
        """
        # Double-check if execution is still marked active. If not, something else
        # might have already called the reset logic, so just ensure reset is done.
        if not self.execution_active:
            self._log("Skipping _on_all_workers_finished actions as execution already marked inactive.")
            # Ensure GUI is reset, just in case. Schedule it to run soon.
            self.master.after(10, self._reset_gui_state)
            return

        self._log("All workers finished processing or stop signal received. Finalizing execution cycle.")
        # Mark the execution cycle as fully stopped/completed.
        self.execution_active = False

        # Check if any tasks remain in the queue (can happen if stop was called aggressively
        # or items were added very late in the process).
        remaining_tasks = self.task_queue.qsize()
        # Determine the final status message based on whether stop was called.
        final_message = "All tasks processed." if not self.stop_event.is_set() else "Execution stopped by user."
        if remaining_tasks > 0:
             # Append information about remaining tasks if any exist.
             final_message += f" ({remaining_tasks} tasks remain in queue)."
             self._log(f"Final state check: {remaining_tasks} tasks remain in queue despite workers finishing.")

        # Log and update the status bar with the final message.
        self._log(f"Final Status: {final_message}")
        self._update_status(final_message)

        # Reset the GUI controls back to their idle state.
        # Schedule this using 'after' to ensure the status update above has a chance to process first.
        self.master.after(50, self._reset_gui_state)

        # Display a pop-up message box informing the user of completion/stoppage.
        info_title = "Execution Complete" if not self.stop_event.is_set() else "Execution Stopped"
        # Schedule the message box as well to ensure it appears after potential GUI resets.
        self.master.after(100, messagebox.showinfo, info_title, final_message)

        self._log("Execution finished cycle. GUI remains open for next run or adding scripts.")


    def _reset_gui_state(self):
        """
        Resets the GUI controls to their default non-executing state.

        Enables/disables buttons and input fields appropriately for the idle state.
        Also ensures the completion lock is released for the next execution cycle.
        """
        try:
            # Check if the main window still exists.
            if not self.master.winfo_exists(): return

            # --- Reset Control States ---
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            # Re-enable list manipulation buttons.
            self.remove_button.config(state=tk.NORMAL)
            self.edit_args_button.config(state=tk.NORMAL)
            # Re-enable adding scripts and related inputs.
            self.add_button.config(state=tk.NORMAL)
            self.duplicates_check.config(state=tk.NORMAL)
            self.args_entry.config(state=tk.NORMAL)
            # Re-enable the parallelism spinbox.
            self.max_parallel_spinbox.config(state=tk.NORMAL)

            self._log("GUI controls reset to idle state.")

        except tk.TclError:
             # Handle error if widgets are destroyed before reset completes.
             self._log("Error resetting GUI state (window likely destroyed).")
        except Exception as e:
             self._log(f"Unexpected error during GUI state reset: {e}")
        finally:
            # --- Release Completion Lock ---
            # Crucially, ensure the completion lock is released, regardless of
            # whether GUI reset succeeded, so the next execution cycle can start.
            try:
                self.completion_lock.release()
                self._log("Completion lock released during GUI reset.")
            except (threading.ThreadError, RuntimeError):
                # Ignore errors if lock was not held or already released.
                pass


    def _on_closing(self):
        """
        Handles the event when the user tries to close the main window.

        If execution is active, it prompts the user for confirmation, explaining
        that running scripts won't be terminated abruptly. If confirmed, it signals
        stop and destroys the window. If execution is not active, it simply
        destroys the window.
        """
        # Check if script execution is currently in progress.
        if self.execution_active:
            # Ask the user to confirm closing while execution is active.
            if messagebox.askyesno("Confirm Exit",
                                   "Execution is active.\n"
                                   "Closing the window now will signal workers to stop picking up *new* tasks, "
                                   "but scripts already running via 'start /wait' will NOT be terminated and will continue in the background.\n\n"
                                   "Stop queuing new tasks and exit the GUI anyway?"):
                # User confirmed exit during execution.
                self._log("Exit requested during active execution. Signaling stop and closing GUI.")
                # Signal workers to stop processing further queue items.
                self.stop_event.set()
                # Send sentinels to potentially wake up workers blocked on the queue.
                for _ in range(len(self.worker_threads) + 1):
                    try: self.task_queue.put(None)
                    except Exception: pass # Ignore errors putting sentinel during shutdown.
                # Optional: Add a very short delay to allow threads to potentially see the stop signal/sentinel.
                # time.sleep(0.1)
                # Destroy the main Tkinter window, ending the application GUI.
                self.master.destroy()
            else:
                # User cancelled the exit request.
                self._log("Exit cancelled by user while execution was active.")
                return # Do not close the window.
        else:
            # Execution is not active, safe to close immediately.
            self._log("Exiting application (execution not active).")
            # Destroy the main Tkinter window.
            self.master.destroy()


# --- Main Execution Block ---
if __name__ == "__main__":
    # This block runs only when the script is executed directly (not imported).

    # Ensure tkinterdnd2 was imported successfully at the top level.
    # The program exits earlier if the import failed.

    # Create the root window using TkinterDnD.Tk to enable drag-and-drop.
    root = TkinterDnD.Tk()

    # Instantiate the main application class, passing the root window.
    app = ScriptExecutorApp(root)

    # Start the Tkinter main event loop. This makes the window visible
    # and responsive to user interactions and events. The program will
    # stay in this loop until the window is closed (e.g., via _on_closing).
    root.mainloop()

    # Log that the mainloop has exited (application is closing).
    print("Application mainloop finished.")