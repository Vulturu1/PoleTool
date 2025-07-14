import customtkinter
import os
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import file_actions as fa


# --- Classes ---

class DraggableLabel(customtkinter.CTkLabel):
    """
    A custom label that can act as a drop target for files.
    """

    def __init__(self, master, app_instance, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_instance  # Store a reference to the main App instance
        # Register the label as a drop target that accepts files.
        self.drop_target_register(DND_FILES)
        # Bind the drop event to the on_drop method.
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.file_path = None

    def on_drop(self, event):
        """
        Handles the file drop event.

        Args:
            event: The event object containing information about the dropped file.
        """
        # The file path is in event.data. It might have curly braces around it.
        path = event.data.strip('{}')

        # Check if the path is a valid file
        if os.path.isfile(path):
            self.file_path = path
            # Get just the filename to display on the label
            filename = os.path.basename(self.file_path)
            self.configure(text=f"Loaded: {filename}")

            # --- CHANGE START: Populate the new filename entry ---
            # Get the filename without its extension
            base_filename = os.path.splitext(filename)[0]
            # Clear the entry and insert the new base filename
            self.app.filename_entry.delete(0, customtkinter.END)
            self.app.filename_entry.insert(0, base_filename)
            # --- CHANGE END ---

            # Enable the button in the main app
            self.app.enable_button()
        else:
            # Handle cases where a folder or invalid item is dropped
            self.configure(text="Invalid drop: Please drop a single file.")
            self.file_path = None
            self.app.disable_button()


class App(TkinterDnD.Tk):
    """
    The main application window which integrates customtkinter and tkinterdnd2.
    """

    def __init__(self):
        super().__init__()

        # --- Basic Window Setup ---
        self.title("Loop Pole Refactor Tool v1.2")
        self.geometry("800x500")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.output_path = ""  # Variable to store the selected output path

        # --- Appearance ---
        customtkinter.set_appearance_mode("Dark")
        customtkinter.set_default_color_theme("blue")

        # --- Main Container Frame ---
        container_frame = customtkinter.CTkFrame(self, corner_radius=15)
        container_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        container_frame.grid_rowconfigure(0, weight=1)
        container_frame.grid_columnconfigure(0, weight=2, minsize=300)
        container_frame.grid_columnconfigure(1, weight=3)

        # --- Left Side (Actions) ---
        actions_frame = customtkinter.CTkFrame(container_frame, fg_color="transparent")
        actions_frame.grid(row=0, column=0, padx=(20, 10), pady=20, sticky="nsew")
        actions_frame.grid_rowconfigure(0, weight=1)
        actions_frame.grid_columnconfigure(0, weight=1)

        scrollable_frame = customtkinter.CTkScrollableFrame(
            actions_frame,
            label_text="File Operations"
        )
        scrollable_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")
        scrollable_frame.grid_columnconfigure(0, weight=1)

        # --- Checkbox Actions with Descriptions ---
        self.checkboxes = {}
        actions_with_descriptions = {
            "Prepare for Vetro": "Refactors file for seamless Vetro import.",
            "Generate MRN": "Generates a formatted Make Ready Sheet.",
            "Generate Verizon Application": "Outputs a Make Ready Sheet for Verizon owned poles.",
            "Generate Frontier Applications (In Development)": "Outputs a folder for Frontier owned poles.",
        }

        row_counter = 0
        for action, description in actions_with_descriptions.items():
            checkbox = customtkinter.CTkCheckBox(scrollable_frame, text=action)
            checkbox.grid(row=row_counter, column=0, padx=10, pady=(10, 0), sticky="w")
            self.checkboxes[action] = checkbox
            row_counter += 1
            desc_label = customtkinter.CTkLabel(
                scrollable_frame, text=description,
                font=customtkinter.CTkFont(size=11), text_color="gray60", justify="left"
            )
            desc_label.grid(row=row_counter, column=0, padx=(38, 10), pady=(0, 10), sticky="nw")
            row_counter += 1

        # --- Right Side (Drop Zone & Controls) ---
        drop_zone_frame = customtkinter.CTkFrame(container_frame, fg_color="transparent")
        drop_zone_frame.grid(row=0, column=1, padx=(10, 20), pady=20, sticky="nsew")
        drop_zone_frame.grid_propagate(False)
        drop_zone_frame.grid_columnconfigure(0, weight=0)
        drop_zone_frame.grid_columnconfigure(1, weight=1)
        drop_zone_frame.grid_rowconfigure(0, weight=1)

        # --- Widgets in Drop Zone Frame (Reordered) ---

        # 1. Drag and Drop Label
        self.drop_label = DraggableLabel(
            master=drop_zone_frame, app_instance=self,
            text="Drag & Drop File Here", font=("Inter", 20),
            text_color="gray", fg_color="#3c3c3c", corner_radius=10
        )
        self.drop_label.grid(row=0, column=0, columnspan=2, padx=0, pady=(0, 10), sticky="nsew")

        # 2. Output Path Selection
        self.output_path_button = customtkinter.CTkButton(
            master=drop_zone_frame, text="Select Output",
            command=self.select_output_path,
            font=("Inter", 12)
        )
        self.output_path_button.grid(row=1, column=0, padx=(0, 5), pady=(10, 10), sticky="w")

        self.output_path_label = customtkinter.CTkLabel(
            master=drop_zone_frame, text="No folder selected...",
            text_color="gray", anchor='w' # Changed anchor to 'w' for consistency
        )
        self.output_path_label.grid(row=1, column=1, padx=(5, 0), pady=(10, 10), sticky="w")

        self.filename_entry = customtkinter.CTkEntry(
            master=drop_zone_frame,
            placeholder_text="(Optional) Output file name..."
        )
        self.filename_entry.grid(row=2, column=0, columnspan=2,  padx=(0, 0), pady=(10, 10), sticky="ew")

        # 4. Process Button (Row is now 3)
        self.process_button = customtkinter.CTkButton(
            master=drop_zone_frame, text="Process File", font=("Inter", 16),
            state="disabled", command=self.process_file_callback,
            corner_radius=10, height=40
        )
        self.process_button.grid(row=3, column=0, columnspan=2, padx=0, pady=(10, 10), sticky="ew")

        # 5. Progress Bar (Row is now 4)
        self.progress_bar = customtkinter.CTkProgressBar(master=drop_zone_frame)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=4, column=0, columnspan=2, padx=0, pady=(10, 5), sticky="ew")

        # 6. Status Label (Row is now 5)
        self.status_label = customtkinter.CTkLabel(master=drop_zone_frame, text="Status: Idle", text_color="gray")
        self.status_label.grid(row=5, column=0, columnspan=2, padx=0, pady=(5, 10), sticky="ew")

    def select_output_path(self):
        """Opens a dialog to select a directory and updates the label."""
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_path = path
            display_path = path
            if len(display_path) > 40:
                display_path = "..." + display_path[-37:]
            self.output_path_label.configure(text=display_path, text_color="white")

    def process_file_callback(self):
        """
        This function is called when the button is clicked.
        It simulates a process and updates the progress bar.
        """
        if not self.drop_label.file_path:
            self.status_label.configure(text="Status: Please drop a file first.")
            return
        if not self.output_path:
            self.status_label.configure(text="Status: Please select an output folder.")
            return

        # --- CHANGE START: Get filename from entry ---
        output_filename = self.filename_entry.get()
        if not output_filename:
            self.status_label.configure(text="Status: Please enter an output file name.")
            return
        # --- CHANGE END ---

        selected_actions = [a for a, c in self.checkboxes.items() if c.get() == 1]

        self.status_label.configure(text="Status: Processing...", text_color="white")
        self.progress_bar.set(0)

        file = fa.read_and_normalize(self.drop_label.file_path)

        file_action_functions = {
            "Prepare for Vetro": lambda: fa.vetro_export(file, self.output_path, output_filename),
            "Generate MRN": lambda: fa.generate_mrn(file, self.output_path, output_filename),
            "Generate Verizon Application": lambda: fa.verizon_app(file, self.output_path, output_filename),
            "Generate Frontier Applications (In Development)": lambda: fa.frontier_pdf(file, self.output_path, output_filename),
        }

        total_steps = len(selected_actions) if selected_actions else 1
        success = True
        for i, action in enumerate(selected_actions or ["Processing..."]):
            self.status_label.configure(text=f"Status: {action}")
            if action in file_action_functions:
                if file_action_functions[action]():
                    success = True
                else:
                    success = False
            self.progress_bar.set((i + 1) / total_steps)
            self.update_idletasks() # Force UI update

        if success:
            self.status_label.configure(text="Status: Complete!", text_color="lightgreen")
        else:
            self.status_label.configure(text="Status: Error!", text_color="red")

    def enable_button(self):
        """Enables the process button."""
        self.process_button.configure(state="normal")

    def disable_button(self):
        """Disables the process button."""
        self.process_button.configure(state="disabled")


# --- Main Execution ---
if __name__ == "__main__":
    app = App()
    app.mainloop()
