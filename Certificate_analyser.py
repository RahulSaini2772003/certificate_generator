from matplotlib.widgets import Button
import os
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import tkinter as tk
from tkinter.simpledialog import Dialog
import pickle
from tkinter import messagebox
from tkinter import filedialog
import subprocess
import csv

class RectangleInputDialog(Dialog):
    def __init__(self, parent, title=None, fonts=None, default_font=None, default_size=None, columns=None):
        self.fonts = fonts
        self.default_font = default_font
        self.default_size = default_size
        self.parent = parent
        self.columns = columns
        super().__init__(parent, title=title)

    def body(self, master):
    
        self.name_label = tk.Label(master, text="Select Column:")
        self.selected_column = tk.StringVar(master)
        self.selected_column.set(self.columns[0])
        self.name_dropdown = tk.OptionMenu(
            master, self.selected_column, *self.columns)
        self.name_dropdown.config(width=15)

    
        self.name_label.grid(row=0, column=0, sticky="w")
        self.name_dropdown.grid(row=0, column=1)

        self.font_label = tk.Label(master, text="Font Style:")
        self.selected_font = tk.StringVar(master, self.default_font)
        self.font_menu = tk.OptionMenu(master, self.selected_font, *self.fonts)
        self.font_label.grid(row=1, column=0, sticky="w")
        self.font_menu.grid(row=1, column=1, columnspan=2, sticky="ew")

        self.custom_size_label = tk.Label(master, text="Font size:")
        self.custom_size_label.grid(row=2, column=0, sticky="w")
        self.custom_size_entry = tk.Entry(master)
        self.custom_size_entry.grid(row=2, column=1, sticky="w", padx=(3, 0))
        self.custom_size_entry.config(width=21)

        self.auto_var = tk.BooleanVar(master, True)
        self.auto_checkbutton = tk.Checkbutton(
            master, text="Auto", variable=self.auto_var, command=self.on_auto_check, height=1, width=4)
        self.auto_checkbutton.grid(row=3, column=1, sticky="w", padx=(0, 5))

        self.thickness_label = tk.Label(master, text="Font Boldness:")
        self.thickness_entry = tk.Entry(master)
        self.thickness_label.grid(row=4, column=0, sticky="w")
        self.thickness_entry.grid(row=4, column=1)
        self.thickness_entry.insert(0, "0")

        self.custom_size_entry.insert(0, "Auto")
        self.custom_size_entry.config(state="disabled")

    def on_auto_check(self):
        if self.auto_var.get():
            self.custom_size_entry.delete(0, tk.END)
            self.custom_size_entry.insert(0, "Auto")
            self.custom_size_entry.config(state="disabled")
        else:
            if self.custom_size_entry.get() == "Auto":
                self.custom_size_entry.config(state="normal")
                self.custom_size_entry.delete(0, tk.END)
            self.custom_size_entry.config(state="normal")

    def apply(self):
        if self.auto_var.get():
            font_size = "Auto"
        else:
            font_size = int(self.custom_size_entry.get())
        self.result = {
            'name': self.selected_column.get(),
            'font_style': self.selected_font.get(),
            'font_size': font_size,
            'thickness': int(self.thickness_entry.get())
        }



def get_font_files(folder):
    try:
        return [f for f in os.listdir(folder) if f.endswith((".ttf", ".otf"))]
    except OSError as e:
        messagebox.showerror("Error", f"Failed to access font folder: {e}")
        return []



def open_files():
    global end_x, end_y, button1_clicked
    button1_clicked = False

    end_x = None
    end_y = None
    try:
        if not all((template_path.get(), csv_path.get(), output_folder_path.get())):
            messagebox.showerror("Error", "Please fill in all fields.")
            return


        def add_text(text):
            text_list.append(text)
            if len(text_list) > 5:
            
                text_list.pop(0)
            update_display()

        def update_display():
            ax.clear()
            ax.imshow(image)
            ax.set_title(
                "Click and drag to select the Text section for printing Text On the Template", pad=15, fontweight='bold', fontsize=14)
            ax.axis("on")

            ax.text(1.12, 0.8, "Terminal :-", ha='center', va='center',
                    fontsize=12, transform=ax.transAxes) 
            for i, text in enumerate(text_list):
          
                y_position = 0.7 - 0.05 * i 

                ax.text(1.08, y_position,
                        f"➤{text}", va='center', fontsize=9, transform=ax.transAxes)

         
            l_mar = 1200
            label1 = ax.text(button1_pos[0] / fig.dpi + button_width / 2 - l_mar,
                             button1_pos[1] / fig.dpi + button_height + 730, 'Generate Sample :-', va='center')
            label2 = ax.text(button2_pos[0] / fig.dpi + button_width / 2 - l_mar,
                             button2_pos[1] / fig.dpi + button_height + 1390, 'View Sample :-', va='center')
            label3 = ax.text(button3_pos[0] / fig.dpi + button_width / 2 - l_mar,
                             button3_pos[1] / fig.dpi + button_height + 2050, 'Reset All :-', va='center')
            fig.canvas.draw()

    
        def button1_click(event):
            global button1_clicked
            if any(isinstance(value, dict) for value in rectangles.values()):
                with open('AppData/rectangles.pkl', 'wb') as file:
                    pickle.dump(rectangles, file)

                try:
                    subprocess.run(
                        ["python", "sample_generator.py"], check=True)
                    button1_clicked = True
                    add_text("Sample generated")
                    print("Template Analysis Successful!!")
                except subprocess.CalledProcessError as e:
                    print(f"Error running {script_path}: {e}")
            else:
                root = tk.Tk()
                root.withdraw()  
                messagebox.showerror(
                    "Error", "Please Analyse Template First By Drag on Template")
                root.mainloop()
        

        def button2_click(event):
            global button1_clicked
            if button1_clicked:
                image_path = "Samples\Sample_Doc.png"
           
                os.startfile(image_path)
                add_text("Viewing Sample")
            else:
                messagebox.showerror("Error", "Please Create Sample to view")

        def button3_click(event, rectangles, start_x, start_y, end_x, end_y, rect):
            global button1_clicked
            rectangles.clear()
            start_x = None
            start_y = None
            end_x = None
            end_y = None
            rect = None
            rectangles['csv_path'] = csv_path_value
            rectangles['template_path'] = temp_path_value
            rectangles['output_folder_path_value'] = output_folder_path_value
            file_path = "Samples\Sample_Doc.png"

            if os.path.exists(file_path):

                try:
                    os.remove(file_path)
                    button1_clicked = False
                except OSError as e:
                    pass
            add_text("Resetting All")
            messagebox.showinfo("Reset Successful",
                                "Reset have been successfully.")


        def update_fontsize(event=None):
            global label1, label2, label3
            button1.label.set_fontsize(
                button1_ax.get_window_extent().height * 0.25)
            button2.label.set_fontsize(
                button2_ax.get_window_extent().height * 0.25)
            button3.label.set_fontsize(
                button3_ax.get_window_extent().height * 0.25)
            if label1:
                label1.set_fontsize(
                    button3_ax.get_window_extent().height * 0.5)
            if label2:
                label2.set_fontsize(
                    button3_ax.get_window_extent().height * 0.5)
            if label3:
                label3.set_fontsize(
                    button3_ax.get_window_extent().height * 0.5)


        def onpress(event):
            nonlocal start_x, start_y
            if event.button == 1:
                start_x = event.xdata
                start_y = event.ydata

        def onmotion(event):
            nonlocal start_x, start_y, rect
            if start_x is not None and start_y is not None:
                end_x = event.xdata
                end_y = event.ydata
                if start_x is not None and start_y is not None and end_x is not None and end_y is not None:
                    width = end_x - start_x
                    height = end_y - start_y
                    if rect is not None:
                        rect.remove()  
                    if start_x is not None and start_y is not None:
                        rect = Rectangle((start_x, start_y), width, height,
                                         linewidth=1, edgecolor='red', facecolor='yellow')
                        ax.add_patch(rect)
                    plt.draw()

        def onrelease(event):
            nonlocal start_x, start_y, rect
            if event.button == 1:  
                end_x = event.xdata
                end_y = event.ydata
                if start_x is not None and start_y is not None:
                    height = end_y - start_y
                    width = end_x - start_x
                if start_x is not None and start_y is not None and end_x is not None and end_y is not None:
                    if rect is not None:
                        rect.remove()  
                    rect = Rectangle((start_x, start_y), width, height,
                                     linewidth=1, edgecolor='red', facecolor='none')
                    ax.add_patch(rect)
                    plt.draw()

                  
                    rectangle_details = None
                    if height > 16.579165515335717 and width > 16.579165515335717:
                        rectangle_details = get_rectangle_details()

                    if rectangle_details is not None:
                        rectangles[rectangle_details['name']] = {'start_x': start_x,
                                                                 'start_y': start_y,
                                                                 'end_x': end_x,
                                                                 'end_y': end_y,
                                                                 'font_style': rectangle_details['font_style'],
                                                                 'font_size': rectangle_details['font_size'],
                                                                 'thickness': rectangle_details['thickness']}

                   
                    start_x = None
                    start_y = None

        def get_rectangle_details():
            global columns
            font_folder = "Fonts"
            try:
                fonts = get_font_files(font_folder)
                dialog = RectangleInputDialog(
                    None, title="Rectangle Details", fonts=fonts, default_font="Default_font.ttf", default_size="Auto", columns=columns)
                return dialog.result
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Error occurred while getting rectangle details: {e}")
                return None
        global label1, label2, label3,columns

        temp_path_value = template_path.get()
        csv_path_value = csv_path.get()
        output_folder_path_value = output_folder_path.get()
        root.destroy()

        with open(csv_path_value, 'r') as csvfile:
            reader = csv.DictReader(csvfile)
            columns = reader.fieldnames

       
        label1 = None
        label2 = None
        label3 = None
       
        button1_pos = (8, 60)
        button2_pos = (8, 40)
        button3_pos = (8, 20)
        button_width = 0.05
        button_height = 0.03

        rectangles = {}
        start_x = None
        start_y = None
        rect = None
        # temp_path_value = "Templates\\templet.jpg"
        image = plt.imread(temp_path_value)

     
        fig, ax = plt.subplots()
        text_list = []
        update_display()

        fig.canvas.mpl_connect('button_press_event', onpress)
        fig.canvas.mpl_connect('motion_notify_event', onmotion)
        fig.canvas.mpl_connect('button_release_event', onrelease)

        button1_pos = (8, 60)
        button2_pos = (8, 40)
        button3_pos = (8, 20)
        button_width = 0.05  
        button_height = 0.03  

   
        button1_ax = plt.axes([button1_pos[0] / fig.dpi + 0.03,
                               button1_pos[1] / fig.dpi, button_width, button_height])
        button1 = Button(button1_ax, 'Sample')
        button1.on_clicked(button1_click)

        button2_ax = plt.axes([button2_pos[0] / fig.dpi + 0.03,
                               button2_pos[1] / fig.dpi, button_width, button_height])
        button2 = Button(button2_ax, 'View')
        button2.on_clicked(button2_click)

        button3_ax = plt.axes([button3_pos[0] / fig.dpi + 0.03,
                               button3_pos[1] / fig.dpi, button_width, button_height])
        button3 = Button(button3_ax, 'Reset')
        button3.on_clicked(lambda event: button3_click(
            event, rectangles, start_x, start_y, end_x, end_y, rect))

   
        update_fontsize()


        fig.canvas.mpl_connect('resize_event', update_fontsize)

        update_display()
       
        rectangles['csv_path'] = csv_path_value
        rectangles['template_path'] = temp_path_value
        rectangles['output_folder_path_value'] = output_folder_path_value

      
        plt.show()

        

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")



def button1_click(event):
    print("Button 1 clicked")


def button2_click(event):
    print("Button 2 clicked")


def button3_click(event):
    print("Button 3 clicked")


def choose_template_path():
    try:
        template_path.set(filedialog.askopenfilename(title="Select Template Path", filetypes=(
            ("All files", "*.*"), ("PNG files", "*.png"), ("JPEG files", "*.jpg"))))
    except Exception as e:
        messagebox.showerror(
            "Error", f"An error occurred while choosing template path: {e}")


def choose_csv_path():
    try:
        csv_path.set(filedialog.askopenfilename(title="Select CSV File",
                                                filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))))
    except Exception as e:
        messagebox.showerror(
            "Error", f"An error occurred while choosing CSV file path: {e}")


def choose_output_folder_path():
    try:
        output_folder_path.set(filedialog.askdirectory(
            title="Select Output Folder"))
    except Exception as e:
        messagebox.showerror(
            "Error", f"An error occurred while choosing output folder path: {e}")


def cancel():
    root.destroy()


columns =[]
root = tk.Tk()
root.title("Certificate Generator")

root.geometry("461x236+638+131")


template_path = tk.StringVar()
csv_path = tk.StringVar()
output_folder_path = tk.StringVar()


welcome_label = tk.Label(
    root, text="Welcome to Certificate Generator", font=("Arial", 14))
welcome_label.grid(row=0, column=0, columnspan=3, padx=(50, 0), pady=20)

template_label = tk.Label(root, text="Template Path:")
template_label.grid(row=1, column=0, padx=5, pady=5)
template_entry = tk.Entry(root, textvariable=template_path, width=40)
template_entry.grid(row=1, column=1, padx=5, pady=5)
template_button = tk.Button(root, text="Choose", command=choose_template_path)
template_button.grid(row=1, column=2, padx=5, pady=5)


csv_label = tk.Label(root, text="CSV File Path:")
csv_label.grid(row=2, column=0, padx=5, pady=5)
csv_entry = tk.Entry(root, textvariable=csv_path, width=40)
csv_entry.grid(row=2, column=1, padx=5, pady=5)
csv_button = tk.Button(root, text="Choose", command=choose_csv_path)
csv_button.grid(row=2, column=2, padx=5, pady=5)

output_folder_label = tk.Label(root, text="Output Folder Path:")
output_folder_label.grid(row=3, column=0, padx=5, pady=5)
output_folder_entry = tk.Entry(
    root, textvariable=output_folder_path, width=40)
output_folder_entry.grid(row=3, column=1, padx=5, pady=5)
output_folder_button = tk.Button(
    root, text="Choose", command=choose_output_folder_path)
output_folder_button.grid(row=3, column=2, padx=5, pady=5)


open_button = tk.Button(root, text="Open", width=0,
                        height=1, command=open_files)
open_button.grid(row=5, column=1, columnspan=10,
                 padx=(170, 80), pady=5, sticky="we")

cancel_button = tk.Button(root, text="Cancel", command=cancel)
cancel_button.grid(row=5, column=2, padx=5, pady=5, sticky="we")

root.mainloop()
