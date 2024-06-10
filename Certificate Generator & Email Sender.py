from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import csv
from email.mime.text import MIMEText
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import pickle
from email import encoders
from PIL import ImageFont, ImageDraw, Image
import warnings
import time
import threading


def load_rectangles(file_path):
    try:
        with open(file_path, 'rb') as file:
            rectangles = pickle.load(file)
        return rectangles
    except FileNotFoundError:
        text_ad("Please Analysis the Certificate Before Generate.")
        print("Error: File not found.")
    except Exception as e:
        text_ad(f"Error loading Data: {e}")
        return {}


def gen_cer_multi(rectangles):
    global font_folder, leader_name_col, members
    try:
        csv_file = rectangles.get('csv_path', '')
        template_path = rectangles.get('template_path', '')
        output_folder = rectangles.get('output_folder_path_value', '')

        if not csv_file or not template_path or not output_folder:
            text_ad("Error: Missing required information.")
            return

        with open(csv_file, 'r') as csvfile:
            reader = csv.DictReader(csvfile)

            headers = reader.fieldnames
            name_index = headers.index(leader_name_col)
            team_member_indices = [name_index]
            for column in members:
                try:
                    index = headers.index(column)
                    team_member_indices.append(index)
                except ValueError:
                    text_ad(f"Column '{column}' not found in the CSV header")
            for row in reader:
                for index in team_member_indices:
                    team_member_name = row[headers[index]]
                    if team_member_name is None:
                        continue
                    print(team_member_indices,team_member_name)
                    with Image.open(template_path) as image:
                        draw = ImageDraw.Draw(image)

                        for col_name, text in row.items():
                            # print(col_name,text , row.items())
                            if col_name in rectangles:
                                # print("etxt" ,text ,'col', col_name,'team',team_member_name ,end='\n')
                                details = rectangles[col_name]
                                start_x, start_y, end_x, end_y = details.get('start_x', 0), details.get(
                                    'start_y', 0), details.get('end_x', 0), details.get('end_y', 0)
                                font_style, font_size, thickness = details.get(
                                    'font_style', ''), details.get('font_size', ''), details.get('thickness', 0)

                                rectangle_width = end_x - start_x
                                rectangle_height = end_y - start_y

                                if font_size == "Auto":
                                    font_size = min(
                                        int(rectangle_height * 0.8), int(rectangle_width * 0.8))
                                else:
                                    font_size = int(font_size)

                                text_position = ((start_x + end_x) / 2,
                                                 (start_y + end_y) / 2)

                                font_file = f"{font_folder}/{font_style}"
                                font = ImageFont.truetype(
                                    font_file, size=font_size)

                                with warnings.catch_warnings():
                                    warnings.simplefilter("ignore")
                                    bbox = draw.textbbox((0, 0), text, font=font)
                                    text_width = bbox[2] - bbox[0]
                                    text_height = bbox[3] - bbox[1]


                                centered_x = text_position[0] - text_width / 2
                                centered_y = text_position[1] - text_height / 2
                                # print(text)

                                draw.text((centered_x, centered_y), text, font=font,
                                          fill=(0, 0, 0), stroke_width=thickness)

                        output_filename = f"{team_member_name.replace(' ', '_')}.png"
                        image.save(f"{output_folder}/{output_filename}")
                        output_filename = output_filename.split(
                            '.')[0].replace('_', ' ')

                        text_ad(
                            f"{output_filename}'s Certificate Generated")
        text_ad(f"Saved at {output_folder}")
        exp_button.config(state='normal')

    except Exception as e:
        text_ad(f"Error generating Certificate: {e}")


def gen_cer_one(text_details):
    try:
        csv_file = text_details.get('csv_path', '')
        template_path = text_details.get('template_path', '')
        output_folder = text_details.get('output_folder_path_value', '')

        if not csv_file or not template_path or not output_folder:
            text_ad("Error: Missing information in Certificate Analysis.")
            return

        with open(csv_file, 'r') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                with Image.open(template_path) as image:
                    draw = ImageDraw.Draw(image)

                    for col_name, text in row.items():
                        if col_name in text_details:
                            details = text_details[col_name]
                            start_x, start_y, end_x, end_y = details.get('start_x', 0), details.get(
                                'start_y', 0), details.get('end_x', 0), details.get('end_y', 0)
                            font_style, font_size, thickness = details.get(
                                'font_style', ''), details.get('font_size', ''), details.get('thickness', 0)

                            rectangle_width = end_x - start_x
                            rectangle_height = end_y - start_y

                            if font_size == "Auto":
                                font_size = min(
                                    int(rectangle_height * 0.8), int(rectangle_width * 0.8))
                            else:
                                font_size = int(font_size)

                            text_position = ((start_x + end_x) / 2,
                                             (start_y + end_y) / 2)

                            font_file = f"{font_folder}/{font_style}"
                            font = ImageFont.truetype(
                                font_file, size=font_size)

                            with warnings.catch_warnings():
                                warnings.simplefilter("ignore")
                               
                                bbox = draw.textbbox((0, 0), text, font=font)
                                text_width = bbox[2] - bbox[0]
                                text_height = bbox[3] - bbox[1]


                            centered_x = text_position[0] - text_width / 2
                            centered_y = text_position[1] - text_height / 2

                            draw.text((centered_x, centered_y), text, font=font,
                                      fill=(0, 0, 0), stroke_width=thickness)

                    output_filename = f"{row.get('Name', '').replace(' ', '_')}.png"

                    image.save(f"{output_folder}/{output_filename}")
                    output_filename = output_filename.split(
                        '.')[0].replace('_', ' ')

                    text_ad(
                        f"{output_filename} Certificate Generated")
        text_ad(f"Saved at {output_folder}")
        exp_button.config(state='normal')
    except Exception as e:
        text_ad(f"Error generating Certificate: {e}")


def for_one(sender_email, password, subject, email_col, name_col, body_msg, rectangles):
    aut_suc = True

    def send_certificate_email(sender_email, password, email_col, certificate_path, subject, body_text, name_col):
        """
        Sends the specified certificate to the recipient's email address.

        Args:
            sender_email (str): Email address of the sender.
            password (str): Password for the sender's email account (store securely, consider environment variables).
            recipient_email (str): Email address of the recipient.
            certificate_path (str): Path to the generated certificate file.
        """
        from email.mime.text import MIMEText
        file_name = certificate_path.split('/')[-1]
        file_name = file_name.split('.')[0].replace("_", " ")
    
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = email_col
        message["Subject"] = subject
        from email.mime.text import MIMEText

        text_part = MIMEText(body_text, "plain")
        message.attach(text_part)

        try:

            with open(certificate_path, "rb") as attachment:
                file_content = attachment.read()
                if file_content: 
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(file_content)
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition",
                                    f"attachment; filename={file_name}")
                    message.attach(part)
                else:
                    print("The file is empty:", certificate_path)
        except FileNotFoundError:
            text_ad("Certificate not found at", certificate_path)
        except PermissionError:
            print("Permission denied for certificate file:", certificate_path)
        except Exception as e:
            text_ad("An error occurred while attaching certificate:", e)

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            try:
                server.login(sender_email, password)
                server.sendmail(sender_email, email_col, message.as_string())
                text_ad(f"{name_col}'s Docs sent successful to {email_col}")
            except smtplib.SMTPAuthenticationError as e:
                nonlocal aut_suc
                aut_suc = False
                text_ad(f"Error: Check your Email id and password: {e}")
            except Exception as e:
                text_ad(f"An error occurred while sending email: {e}")

    csv_file = rectangles['csv_path']
    output_folder = rectangles['output_folder_path_value']

    with open(csv_file, 'r') as csvfile:
        reader = csv.reader(csvfile)
        data = list(reader) 
        headers = data[0] 

    name_index = headers.index(f"{name_col}")
    email_index = headers.index(f"{email_col}")

    for row in data[1:]:
        if not aut_suc:
            text_ad("Authentication failed. Stopping the email sending process.")
            break
        name = row[name_index]
        email = row[email_index]
        certificate_path = f"{output_folder}/{name.replace(' ','_')}.png"

        body_text = f'''Dear {name},

    {body_msg}'''

        text_ad("Sending emails, please wait......")
        send_certificate_email(sender_email, password, email,
                               certificate_path, subject, body_text, name)
    if aut_suc:
        text_ad("All Certificates Sent succeed")


def for_multi(sender_email, password, subject, email_col, leader_name_col, members, body_msg, rectangles):
    aut_suc_multi = True

    def send_certificate_email(sender_email, password, recipient_email, subject, valid_team_members, body_msg, output_folder):
        """
        Sends the specified certificates to the recipient's email address.

        Args:
            sender_email (str): Email address of the sender.
            password (str): Password for the sender's email account.
            recipient_email (str): Email address of the recipient.
            certificate_paths (list): List of paths to the certificate files.
            subject (str): Subject of the email.
            body_text (str): Body text of the email.
        """

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = recipient_email
        message["Subject"] = subject

        body_text = f'''Dear {valid_team_members[0]},

    {body_msg}'''

        message.attach(MIMEText(body_text, "plain"))

        try:
           
            for member in valid_team_members:
                certificate_path = f"{output_folder}/{member.replace(' ','_')}.png"
                file_name = certificate_path.split('/')[-1]
                file_name = file_name.split('.')[0].replace("_", " ")
                with open(certificate_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition",
                                    f"attachment; filename={file_name}")
                    message.attach(part)
        except FileNotFoundError as re:
            text_ad("Certificate not found at", certificate_path,
                    "So, Email not sent successful on", recipient_email)
            return
        except PermissionError:
            print("Permission denied for certificate file:", certificate_path)
            return
        except Exception as e:
            print("An error occurred while attaching certificate:", e)
            return

    
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls() 
            try:
                server.login(sender_email, password)
                server.sendmail(sender_email, recipient_email,
                                message.as_string())
            except smtplib.SMTPAuthenticationError as e:
                nonlocal aut_suc_multi
                aut_suc_multi = False
                text_ad(f"Error: Check your Email id and password: {e}")
                return
            except Exception as e:
                text_ad(f"An error occurred while sending email: {e}")
        text_ad(
            f"{valid_team_members}'s Docs sent successful to {recipient_email}")

    csv_file = rectangles['csv_path']
    output_folder = rectangles['output_folder_path_value']

    text_ad("Sending emails, please wait......")
    with open(csv_file, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        headers = reader.fieldnames 

        for row in reader:
            if not aut_suc_multi:
                text_ad("Authentication failed. Stopping the email sending process.")
                break
            try:
                name = row[f"{leader_name_col}"]
                member_data = [name] + [row[member]
                                        for member in members if member in row]

                recipient_email = row[email_col]

                valid_team_members = [
                    member for member in member_data if member is not None]
                valid_team_members = [member.strip().title(
                ) for member in valid_team_members if member.strip()]
                send_certificate_email(sender_email, password, recipient_email,
                                       subject, valid_team_members, body_msg, output_folder)
            except Exception as e:
                text_ad(
                    f"An error occurred for Team Leader's Email {recipient_email}: {e}. Skipping to the next Team.")
                continue
    if aut_suc_multi:
        text_ad("All Certificates Sent succeed")


def export():
    global terminal
    if not terminal:
        messagebox.showwarning("No Data", "Generate Certificates to Export.")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt", filetypes=[("Text Files", "*.txt")],initialfile='Genereted Certificates')
    if file_path:
        with open(file_path, 'w') as file:
            for item in terminal:
                file.write(f"{item}\n")
        terminal.clear()
        update_terminal_display()
        exp_button.config(state='disabled')
        messagebox.showinfo("Saved", "File saved successfully!")


def update_terminal_display():
    global terminal_label, terminal
    max_len = 100
    ter_l = ttk.Label(main_frame, text="Terminal :- \n",
                      font=("Helvetica", 15), wraplength=220)
    ter_l.grid(row=2, column=4, padx=30, sticky="w")
    terminal_text = "\n".join(f"      ➤ {text[:max_len]}...\n" if len(
        text) > max_len else f"      ➤ {text}\n" for text in terminal)
    num_lines = len(terminal_text.split("\n"))
    pady = max(0, (30 - num_lines) * 14)
    new_terminal_label = ttk.Label(
        main_frame, text=terminal_text, font=("Helvetica", 8), wraplength=320)
    new_terminal_label.grid(row=3, column=4, padx=40,
                            pady=(0, pady+50), sticky="w", rowspan=11)
    if terminal_label:
        terminal_label.grid_forget()
    terminal_label = new_terminal_label
    root.update()


def on_select(value):
    pass  


def on_add():
    global selected_columns, av_colu
    selected_value = selected_column.get()
    if selected_value == "----   Select Column   ----":
        return
    if selected_value not in selected_columns:
        selected_columns.append(f"'{selected_value}'")
        if selected_value in av_colu:
            index = av_colu.index(selected_value)
           
            if index < len(av_colu):
               
                dropdown['menu'].delete(index)
                del av_colu[index]
            else:
                print("Index out of range. Cannot delete.")
        members_entry.set(", ".join(selected_columns))
        selected_column.set("----   Select Column   ----")
    else:
        print("Selected column already in selected_columns list or empty")


def switch_callback():
    global is_on

    if is_on:
        switch.config(image=off)
        is_on = False

        input_field.config(state="disabled")
        dropdown.config(state="disabled")
        add_button.config(state="disabled")
    else:
        switch.config(image=on)
        is_on = True
        input_field.config(state="normal")
        dropdown.config(state="normal")
        add_button.config(state="normal")


def text_ad(text):
    global terminal
    terminal.append(text)
    if len(terminal) >= 15:
        terminal.pop(0)
    update_terminal_display()


def generate():
    global rectangles, is_on, leader_name_col, sel_col, flag, terminal, members
    terminal.clear()
    if leader_name_col_entry.get() != sel_col:
        leader_name_col = leader_name_col_entry.get()
    if not leader_name_col:
        messagebox.showwarning(
            "Warning", "Please fill above all the required fields.")
        return
    if is_on:
        if (is_on and not members_entry.get()):
            messagebox.showwarning(
                "Warning", "Please fill in all the required fields.")
            return
        members = [item.strip().strip("'")
                   for item in members_entry.get().split(',')]
        if show_confirmation_dialog("Are you sure you have chosen all members columns?"):
            if rectangles:
                reset_entry_fields()
                if is_on:
                    gen_cer_multi(rectangles)
                else:
                    gen_cer_one(rectangles)
                flag = False
            else:
                text_ad("Certificate Analysis Data Not found")
        else:
            return
    else:
        if show_confirmation_dialog("Yes, If You Don't Have Team Members"):
            if rectangles:
                reset_entry_fields()
                if is_on:
                    gen_cer_multi(rectangles)
                else:
                    gen_cer_one(rectangles)
                flag = False
            else:
                text_ad("Certificate Analysis Data Not found")
        else:
            return


def show_confirmation_dialog(msg):
    return messagebox.askyesno("Confirmation", msg)


def exp():
    global terminal
    if not terminal:
        messagebox.showwarning("No Data", "Generate Certificates to Export.")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt", filetypes=[("Text Files", "*.txt")],initialfile='Email Details')
    if file_path:
        terminal_without_waiting = [
            item for item in terminal if "Sending emails, please wait......" not in item]
        with open(file_path, 'w') as file:
            for item in terminal_without_waiting:
                file.write(f"{item}\n")
        terminal.clear()
        update_terminal_display()
        export_button.config(state='disabled')
        messagebox.showinfo("Saved", "File saved successfully!")


def send_emails_in_thread():
    submit_details()


def sentt():
    global terminal
    terminal.clear()
    update_terminal_display()
    threading.Thread(target=send_emails_in_thread).start()


def reset_entry_fields():
    sender_email_entry.delete(0, 'end')
    password_entry.delete(0, 'end')
    subject_entry.delete(0, 'end')
    body_msg_entry.delete(1.0, 'end')
    leader_name_col_entry.set(sel_col)
    email_col_entry.set(sel_col)
    selected_column.set("----   Select Column   ----")
    members_entry.set("")


def submit_details():
    email_col = ""
    global leader_name_col, members, sel_col, flag
    if flag:
        messagebox.showwarning(
            "Warning", "Please Generate Certificates For Sent.")
        return
    sender_email = sender_email_entry.get().lower()
    password = password_entry.get().lower()
    subject = subject_entry.get().title()
    if email_col_entry.get() != sel_col:
        email_col = email_col_entry.get()
    leader_name_col = "Name"

    body_msg = body_msg_entry.get("1.0", "end-1c").strip()

    if not all([sender_email, password, subject, email_col, body_msg]):
        messagebox.showwarning(
            "Warning", "Please fill in all the required fields.")
        return
    if "@" not in sender_email:
        messagebox.showwarning("Warning", "Invalid Email.")
        return

    reset_entry_fields()
    if is_on == False:
        for_one(sender_email, password, subject,
                email_col, leader_name_col, body_msg, rectangles)
        export_button.config(state='normal')
    else:
        for_multi(sender_email, password, subject, email_col,
                  leader_name_col, members, body_msg, rectangles)
        export_button.config(state='normal')


members = []
flag = True
leader_name_col = ""
is_on = False
terminal = []
selected_columns = []
rectangles = load_rectangles("AppData/rectangles.pkl")
font_folder = "Fonts"
root = Tk()
root.title("Email Sender")
root.geometry("1000x665+256+54")
root.resizable(False, False)
widt = 50
height = 4
entry_font = ("Helvetica", 13)
l_font = ("Helvetica", 10)
pad = (0, 15)
sel_col = "---- Select Column ----"

csv_file = rectangles.get('csv_path', '')
with open(csv_file, 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    csv_columns = reader.fieldnames
av_colu = csv_columns


main_frame = ttk.Frame(root, padding=(10, 10))
main_frame.pack(fill=BOTH, expand=True)

title_label = ttk.Label(main_frame, text="Certificate Generator & Email Sender",
                        font=("Helvetica", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=2, pady=(3, 5), padx=(130, 0))

slogan_label = ttk.Label(
    main_frame, text="Empower Your Reach: Send Smarter, Send Faster with Email Sender in Bulk!", font=("Helvetica", 8))
slogan_label.grid(row=1, column=0, columnspan=2, pady=(0, 30), padx=(130, 0))

leader_name_col_label = ttk.Label(
    main_frame, text="Leader/Name Column:", font=l_font)
leader_name_col_label.grid(row=2, column=0, padx=5, sticky="w")
leader_name_col_entry = StringVar()
leader_name_col_entry.set(sel_col)
dropdown_menu = OptionMenu(main_frame, leader_name_col_entry, *csv_columns)
dropdown_menu.config(width=widt+9, font=("Helvetica", 9))
dropdown_menu.grid(row=2, column=1, padx=5, pady=pad, sticky="w")

input_field_label = ttk.Label(
    main_frame, text="Team Members:", font=l_font)
input_field_label.grid(row=3, column=0, padx=5, sticky="w", pady=(0, 15))

on = PhotoImage(file="Imgs/on.png")
off = PhotoImage(file="Imgs/off.png")
switch = Button(main_frame, image=off,
                command=switch_callback, relief="flat", bd=0)
switch.grid(row=3, column=1, columnspan=1, padx=(0, 410), pady=(0, 16))

members_entry = StringVar()
input_field = ttk.Entry(
    main_frame, textvariable=members_entry, width=widt-6, font=entry_font)
input_field.config(state=DISABLED)
input_field.grid(row=3, column=1, padx=(60, 0), pady=pad, sticky="w")



label = ttk.Label(main_frame, text="Select Column:", font=l_font)
label.grid(row=4, column=0, padx=5, pady=pad, sticky="w")

selected_column = StringVar()
selected_column.set("----   Select Column   ----")  
dropdown = OptionMenu(main_frame, selected_column, *av_colu)
dropdown.config(width=widt-4, font=("Helvetica", 9),
                state=DISABLED)
dropdown.grid(row=4, column=1, padx=5, pady=pad, sticky="w")

add_button = ttk.Button(main_frame, text="Add", command=on_add)
add_button.config(state=DISABLED)
add_button.grid(row=4, column=1, columnspan=1, pady=pad, padx=(380, 0))

submit_button = ttk.Button(main_frame, text="Generate", command=generate)
submit_button.grid(row=5, column=0, columnspan=2, pady=(0, 35), padx=(0, 100))

exp_button = ttk.Button(main_frame, text="Export to File", command=export)
exp_button.grid(row=5, column=0, columnspan=2, pady=(0, 35), padx=(330, 0))
exp_button.config(state='disabled')


sender_email_label = ttk.Label(main_frame, text="Sender Email:", font=l_font)
sender_email_label.grid(row=6, column=0, padx=5, sticky="w")
sender_email_entry = ttk.Entry(main_frame, width=widt, font=entry_font)
sender_email_entry.grid(row=6, column=1, padx=5, pady=pad, sticky="w")


password_label = ttk.Label(main_frame, text="Password:", font=l_font)
password_label.grid(row=7, column=0, padx=5, sticky="w")
password_entry = ttk.Entry(main_frame, show="*", width=widt, font=entry_font)
password_entry.grid(row=7, column=1, padx=5, pady=pad, sticky="w")

email_col_label = ttk.Label(main_frame, text="Email Column:", font=l_font)
email_col_label.grid(row=8, column=0, padx=5, sticky="w")
email_col_entry = StringVar()
email_col_entry.set(sel_col)  
dropdown_menu = OptionMenu(main_frame, email_col_entry, *csv_columns)
dropdown_menu.config(width=widt+9, font=("Helvetica", 9))
dropdown_menu.grid(row=8, column=1, padx=5, pady=pad, sticky="w")



subject_label = ttk.Label(main_frame, text="Subject:", font=l_font)
subject_label.grid(row=9, column=0, padx=5, sticky="w")
subject_entry = ttk.Entry(main_frame, width=widt, font=entry_font)
subject_entry.grid(row=9, column=1, padx=5, pady=pad, sticky="w")


body_msg_label = ttk.Label(main_frame, text="Body Message:", font=l_font)
body_msg_label.grid(row=10, column=0, padx=5, sticky="w")
body_msg_entry = Text(main_frame, height=8,
                      font=("Helvetica", 10), width=widt+15)
body_msg_entry.grid(row=10, column=1, padx=5, pady=pad, sticky="w")


submit_button = ttk.Button(main_frame, text="Sent", command=sentt)
submit_button.grid(row=11, column=0, columnspan=2, pady=(0, 35), padx=(0, 90))

export_button = ttk.Button(main_frame, text="Export to File", command=exp)
export_button.grid(row=11, column=0, columnspan=2, pady=(0, 35), padx=(330, 0))
export_button.config(state='disabled')


terminal_label = ttk.Label(
    main_frame, text="Terminal :-\n", font=l_font, wraplength=400)
terminal_label.grid(row=3, column=4, padx=40, pady=pad, sticky="w")


update_terminal_display()

root.mainloop()
