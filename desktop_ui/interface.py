"""
This module defines the AudioInterface class for audio analysis.

It provides a graphical user interface for selecting audio files and configuring analysis parameters.

"""
import json
import os.path
import config
from tkinter import Tk, filedialog, messagebox, Text, END, IntVar, StringVar, BooleanVar
from tkinter import ttk
from threading import Thread

from signal_processing.signals import get_signal_sets
from export import export


class AudioInterface:
    tasks = []
    ref_files_list_label = None
    c_files_list_label = None
    logbox = None

    def __init__(self):
        """
        Initialize the AudioInterface class.
        """
        self.root = Tk()

        # Load default data
        export_config = config.ExportConfig()

        self.time_fractions: IntVar = IntVar(value=export_config.time_fractions)
        self.c_name: StringVar = StringVar(value=export_config.c_name)
        self.ref_name: StringVar = StringVar(value=export_config.ref_name)
        self.export_audio_checkbox = BooleanVar(value=export_config.export_audio)
        self.export_fft_checkbox = BooleanVar(value=export_config.export_fft)
        self.regions = export_config.frequency_regions
        self.c_files = export_config.c_files
        self.ref_files = export_config.ref_files

        self.setup_ui()

    def setup_ui(self):
        """
        Setup the user interface. Add frames, buttons, labels, etc.
        """
        self.root.title("audio analysis")
        self.root.resizable(False, False)

        frm = ttk.Frame(self.root, padding=10)
        frm.grid()

        # Setup frequency ranges input
        input_frame = self.get_input_frame(self.root)
        input_frame.grid(column=3, row=0, rowspan=3, sticky='nsew')

        time_fractions_frame = self.get_time_fractions_frame(frm)
        time_fractions_frame.grid(column=2, row=1)

        ttk.Button(frm, text="Select Files For C", command=lambda: self.open_files('C')) \
            .grid(column=0, row=0, sticky='nsew')
        ttk.Button(frm, text="Select Files For REF", command=lambda: self.open_files('REF')) \
            .grid(column=1, row=0, sticky='nsew')
        self.c_files_list_label = ttk.Label(frm, text="Selected C files will show up here")
        self.c_files_list_label.grid(column=0, row=1)
        self.ref_files_list_label = ttk.Label(frm, text="Selected REF files will show up here")
        self.ref_files_list_label.grid(column=1, row=1)
        self.update_files('c', self.c_files)
        self.update_files('ref', self.ref_files)

        ttk.Button(frm, text="Run analysis", command=self.run_analysis).grid(column=2, row=0, sticky='nsew')

        self.logbox = Text(frm)
        self.logbox.insert(END, "Logs will show up here\n")
        self.logbox.config(state='normal')
        self.logbox.grid(column=0, columnspan=3, row=2)

        scrollbar = ttk.Scrollbar(frm, command=self.logbox.yview)
        scrollbar.grid(column=3, row=2, sticky='nsew')

        self.logbox['yscrollcommand'] = scrollbar.set

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.mainloop()

    def get_input_frame(self, parent) -> ttk.Frame:
        """
        Create a frame for inputting frequency ranges.
        """
        main_frame = ttk.Frame(parent)
        entry_frame = ttk.Frame(main_frame)
        ttk.Label(main_frame, text='Enter frequency ranges:').pack(pady=10, padx=10)
        entry_frame.pack(padx=10)

        start_entry = ttk.Entry(entry_frame, width=6)
        end_entry = ttk.Entry(entry_frame, width=6)
        start_entry.pack(side='left')
        end_entry.pack(side='left')

        def add_range(start, end):
            region = (start, end)
            # Add current widget in the main frame
            ranges_frame = ttk.Frame(main_frame)
            ranges_frame.pack(padx=10)

            # Create label and functional remove button for the widget
            ttk.Label(ranges_frame, text=f'{start}Hz to {end}Hz').pack(side='left')
            ttk.Button(ranges_frame, text='Remove',
                       command=lambda: remove_range(ranges_frame, region)).pack(side='left')

            if region not in self.regions:
                self.regions.append(region)

        def remove_range(ranges_frame, region_value):
            AudioInterface.clear_frame(ranges_frame)
            self.regions.remove(region_value)

        def add_range_callback():
            # Validation
            try:
                start = int(start_entry.get())
                end = int(end_entry.get())
            except ValueError:
                messagebox.showerror(title='cannot add range', message='frequencies must be integers')
                return

            if end <= start:
                messagebox.showerror(title='cannot add range',
                                     message='end frequency must be more than start frequency')
                return

            if (start, end) in self.regions:
                messagebox.showerror(title='cannot add range',
                                     message='this range is already added')
                return

            # Actual command for adding ranges
            add_range(start, end)

        ttk.Button(entry_frame, text='Add', width=3, command=add_range_callback).pack(side='left')

        # Add default ranges to this frame at the start
        [add_range(region[0], region[1]) for region in self.regions]

        return main_frame

    def get_time_fractions_frame(self, parent) -> ttk.Frame:
        """
        Create a frame for inputting time fractions.
        """
        main_frame = ttk.Frame(parent)

        def show_input_frame(parent_frame, label, variable):
            frame = ttk.Frame(parent_frame)
            frame.pack(side='top')
            ttk.Label(frame, text=label, width=14).pack(side='left')
            ttk.Entry(frame, width=6, textvariable=variable).pack(side='left')

        show_input_frame(main_frame, 'Time Fractions: ', self.time_fractions)
        show_input_frame(main_frame, 'C Column Name: ', self.c_name)
        show_input_frame(main_frame, 'Ref Column Name: ', self.ref_name)
        ttk.Checkbutton(main_frame, text='Export .wav Fractions', variable=self.export_audio_checkbox) \
            .pack(side='top')
        ttk.Checkbutton(main_frame, text='Export FFT Graphs', variable=self.export_fft_checkbox) \
            .pack(side='top')

        return main_frame

    def on_close(self):
        """
        Save current data to a json file and close the window.
        """
        try:
            export_config = config.ExportConfig(c_files=self.c_files, ref_files=self.ref_files,
                                                c_name=self.c_name.get(), ref_name=self.ref_name.get(),
                                                time_fractions=self.time_fractions.get(),
                                                export_audio=self.export_audio_checkbox.get(),
                                                export_fft=self.export_fft_checkbox.get(),
                                                frequency_regions=self.regions, json_load_path=None)
            json_data = export_config.get_json()
            with open('selected_files.json', 'w') as file:
                file.write(json.dumps(json_data))
        except Exception as e:
            messagebox.showerror(title='could not save data', message=str(e))

        for task in self.tasks:
            task.cancel()
        self.root.destroy()

    def open_files(self, files_category):
        """
        Open a file dialog for selecting files.
        """
        res = filedialog.askopenfiles(mode='r', title='Select Files For ' + files_category)
        files = sorted([x.name for x in res])
        self.update_files(str.lower(files_category), files)

    def update_files(self, files_category, files):
        """
        Update the list of selected files.
        """
        self.__setattr__(files_category + '_files', files)
        self.display_files(files_category)

    def display_files(self, files_category):
        """
        Display the list of selected files.
        """
        attr_name = files_category + '_files_list_label'
        selected_files = self.__getattribute__(files_category + '_files')
        if len(selected_files) > 0:
            files_list = f'{files_category} directory: \n  ' \
                         f'{os.path.dirname(selected_files[0])}\n{files_category} files:\n  '
        else:
            files_list = ''
        files_list += '\n  '.join([os.path.basename(x) for x in selected_files])
        self.__getattribute__(attr_name).config(text=files_list)

    def run_analysis(self):
        """
        Run the analysis based on the current data.
        """
        if len(self.c_files) == 0 or len(self.ref_files) == 0:
            messagebox.showerror(title='cannot export', message='files not selected')
            return
        if len(self.c_files) != len(self.ref_files):
            messagebox.showerror(title='cannot export', message='number of files must be equal')
            return
        try:
            get_signal_sets(len(self.c_files))
        except Exception as e:
            messagebox.showerror(title='cannot export', message=str(e))
            return

        # Create export config using current data
        export_config = config.ExportConfig(c_files=self.c_files, ref_files=self.ref_files,
                                            c_name=self.c_name.get(), ref_name=self.ref_name.get(),
                                            time_fractions=self.time_fractions.get(),
                                            export_audio=self.export_audio_checkbox.get(),
                                            export_fft=self.export_fft_checkbox.get(),
                                            frequency_regions=self.regions, json_load_path=None)

        thread = Thread(target=export.create_export, args=(export_config, self.log), daemon=True)

        thread.start()

    # Modified version of https://stackoverflow.com/a/60034559
    @staticmethod
    def clear_frame(frame):
        """
        Static method for clearing all widgets in a frame.
        """
        for widget in frame.winfo_children():
            widget.destroy()
        frame.pack_forget()

    def log(self, text, end='\n'):
        """
        Log a message to the logbox.
        """
        print(text, end=end)
        self.logbox.insert(END, text + end)
        self.logbox.yview(END)
