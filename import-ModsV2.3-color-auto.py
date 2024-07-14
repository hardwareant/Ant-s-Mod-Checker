import os
import csv
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
import webbrowser
from pathlib import Path
import winreg
import string

# Function to get the Steam installation path dynamically
def get_steam_path_from_env():
    steam_path = os.getenv('STEAMPATH') or os.getenv('ProgramFiles(x86)')
    if steam_path:
        potential_path = Path(steam_path) / "Steam"
        if (potential_path / "steam.exe").exists():
            return potential_path
    return None

def get_steam_install_path():
    env_path = get_steam_path_from_env()
    if env_path:
        return env_path

    # Search for Steam across all available drives
    drives = [f"{char}:" for char in string.ascii_uppercase if Path(f"{char}:/").exists()]
    likely_folders = ["Program Files (x86)/Steam", "Program Files/Steam", "Steam"]

    for drive in drives:
        for folder in likely_folders:
            check_path = Path(drive) / folder
            if (check_path / "steam.exe").exists():
                return check_path

    # Check the registry as the last resort
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Valve\Steam")
        install_path, _ = winreg.QueryValueEx(key, "InstallPath")
        steam_path = Path(install_path)
        if (steam_path / "steam.exe").exists():
            return steam_path
    except Exception as e:
        print(f"Error accessing registry: {e}")

    return None

# Function to retrieve steam library paths
def get_steam_library_paths(steam_path):
    library_folders_path = steam_path / "steamapps/libraryfolders.vdf"
    library_paths = [steam_path / "steamapps"]
    try:
        with open(library_folders_path, 'r') as file:
            lines = file.readlines()
            for line in lines:
                if '"path"' in line:
                    path = line.split('"')[-2]
                    library_paths.append(Path(path) / "steamapps")
    except Exception as e:
        print(f"Error reading {library_folders_path}: {e}")
    return library_paths

# Function to find Project Zomboid workshop
def find_project_zomboid_workshop():
    steam_path = get_steam_install_path()
    if not steam_path:
        return None, None
    
    library_paths = get_steam_library_paths(steam_path)
    workshop_directory = None
    zomboid_directory = None
    
    for library_path in library_paths:
        project_zomboid_path = library_path / "common/Project Zomboid"
        if not project_zomboid_path.exists():
            project_zomboid_path = library_path / "common/ProjectZomboid"
        if project_zomboid_path.exists():
            zomboid_directory = project_zomboid_path
            workshop_path = library_path / "workshop/content/108600"
            if workshop_path.exists():
                workshop_directory = workshop_path
                break

    return workshop_directory, zomboid_directory

# Function to read mod information
def read_mod_info(mod_info_path):
    mod_id = mod_name = None
    try:
        with open(mod_info_path, 'r') as file:
            for line in file:
                if line.startswith('id='):
                    mod_id = line.strip().split('=')[1]
                elif line.startswith('name='):
                    mod_name = line.strip().split('=')[1]
                if mod_id and mod_name:
                    break
    except Exception as e:
        print(f"Error reading {mod_info_path}: {e}")
    return mod_id, mod_name

# Function to generate lists of mods
def generate_lists(directory):
    workshop_items = []
    mods = []
    display_items = []
    
    for gameid in os.listdir(directory):
        gameid_path = os.path.join(directory, gameid)
        
        if os.path.isdir(gameid_path):
            mods_path = os.path.join(gameid_path, 'mods')
            
            if os.path.exists(mods_path):
                for mod in os.listdir(mods_path):
                    mod_path = os.path.join(mods_path, mod)
                    
                    if os.path.isdir(mod_path):
                        mod_info_path = os.path.join(mod_path, 'mod.info')
                        
                        if os.path.exists(mod_info_path):
                            mod_id, mod_name = read_mod_info(mod_info_path)
                            
                            if mod_id and mod_name:
                                workshop_items.append(gameid)
                                mods.append(mod_name)
                                display_items.append((mod_name, gameid))
                            else:
                                print(f"No 'id' or 'name' found in {mod_info_path}")
                        else:
                            print(f"No 'mod.info' found in {mod_path}")
    return workshop_items, mods, display_items

# Function to save data to files
def save_to_files():
    # Get the user's Documents folder
    documents_folder = Path.home() / "Documents"
    
    # Create the Zomboid-ModInfo folder
    save_folder = documents_folder / "Zomboid-ModInfo"
    save_folder.mkdir(parents=True, exist_ok=True)

    csv_file = save_folder / "ModInfo.csv"
    with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(['WorkshopItems', 'Mods'])
        for i in range(len(workshop_items)):
            writer.writerow([workshop_items[i], mods[i]])
    print(f'Data successfully written to {csv_file}')
    
    workshop_items_file = save_folder / "WorkshopItems.txt"
    with open(workshop_items_file, 'w', encoding='utf-8') as file:
        file.write(';'.join(workshop_items))
    print(f'WorkshopItems successfully written to {workshop_items_file}')
    
    mods_file = save_folder / "Mods.txt"
    with open(mods_file, 'w', encoding='utf-8') as file:
        file.write(';'.join(mods))
    print(f'Mods successfully written to {mods_file}')
    
    urls_file = save_folder / "ModURLs.txt"
    url_list = [f"https://steamcommunity.com/sharedfiles/filedetails/?id={id_}" for id_ in workshop_items]
    with open(urls_file, 'w', encoding='utf-8') as file:
        file.write('\n'.join(url_list))
    print(f'URLs successfully written to {urls_file}')

    messagebox.showinfo("Success", f"Data successfully saved to {save_folder}!")

# Function to insert clickable links
def insert_link_text(text_widget, display_text, url, tag_name):
    text_widget.insert(tk.END, display_text)
    text_widget.tag_add(tag_name, "end-%dc" % (len(display_text) + 1), "end-%dc" % (1))
    text_widget.tag_bind(tag_name, "<Button-1>", lambda e, url=url: webbrowser.open_new(url))
    text_widget.tag_config(tag_name, foreground="blue", underline=True)

# Function to select a directory
def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        process_directory(directory)

# Function to process the selected directory
def process_directory(directory):
    global workshop_items, mods, display_items
    directory_label.config(text=f"Selected Directory: {directory}")
    workshop_items, mods, display_items = generate_lists(directory)
    
    # Update mod count
    mod_count_label.config(text=f"Number of Mods: {len(mods)}")
    
    workshop_items_text.config(state=tk.NORMAL)
    mods_text.config(state=tk.NORMAL)
    urls_text.config(state=tk.NORMAL)
    workshop_items_text.delete(1.0, tk.END)
    mods_text.delete(1.0, tk.END)
    urls_text.delete(1.0, tk.END)
    workshop_items_text.insert(tk.END, ';'.join(workshop_items))
    mods_text.insert(tk.END, ';'.join(mods))
    for mod_name, gameid in display_items:
        display_text = f"{mod_name} - {gameid}\n"
        url = f"https://steamcommunity.com/sharedfiles/filedetails/?id={gameid}"
        insert_link_text(urls_text, display_text, url, mod_name)
    workshop_items_text.config(state=tk.DISABLED)
    mods_text.config(state=tk.DISABLED)
    urls_text.config(state=tk.DISABLED)

# Function to clear text fields
def clear_text():
    workshop_items_text.config(state=tk.NORMAL)
    mods_text.config(state=tk.NORMAL)
    urls_text.config(state=tk.NORMAL)
    workshop_items_text.delete(1.0, tk.END)
    mods_text.delete(1.0, tk.END)
    urls_text.delete(1.0, tk.END)
    workshop_items_text.config(state=tk.DISABLED)
    mods_text.config(state=tk.DISABLED)
    urls_text.config(state=tk.DISABLED)

    # Clear mod count
    mod_count_label.config(text="Number of Mods: 0")

# Function to exit the application
def exit_app():
    root.quit()

# TKinter GUI setup
root = tk.Tk()
root.title("Ant's Mod-Checker")
root.geometry("820x620")

# Styling
style = {
    "bg": "#212121",  # Dark gray background
    "frame_bg": "#282828",  # Slightly lighter gray for frames
    "button_bg": "#323232",  # Darker buttons
    "button_active_bg": "#424242",  # Slightly lighter when active
    "highlight_bg": "#ffffff",  # White background for text areas
    "font": ("Arial", 12, "bold"),  # Bold font style
    "label_font": ("Arial", 10, "bold"),  # Bold font for labels
    "text_fg": "#000000",  # Black text color for text areas
    "button_fg": "#d32f2f",  # Red font color for buttons
    "label_fg": "#d32f2f",  # Red font color for labels
    "frame_highlight": "#d32f2f"  # Dark red frame highlight color
}

root.configure(bg=style["bg"])

frame = tk.Frame(root, padx=20, pady=20, bg=style["frame_bg"], highlightbackground=style["frame_highlight"], highlightthickness=2)
frame.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

label = tk.Label(frame, text="Select Project Zomboid Mods Directory:", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
label.pack(pady=5)

select_button = tk.Button(frame, text="Select Directory", command=select_directory, bg=style["button_bg"], activebackground=style["button_active_bg"], font=style["font"], fg=style["button_fg"])
select_button.pack(pady=5)

steam_dir_label = tk.Label(frame, text="Steam Installation Directory: Not Found", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
steam_dir_label.pack(pady=5)

pzm_dir_label = tk.Label(frame, text="Project Zomboid Directory: Not Found", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
pzm_dir_label.pack(pady=5)

directory_label = tk.Label(frame, text="", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
directory_label.pack(pady=5)

# Label for mod count
mod_count_label = tk.Label(frame, text="Number of Mods: 0", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
mod_count_label.pack(pady=5)

content_frame = tk.Frame(frame, bg=style["frame_bg"])
content_frame.pack(expand=True, fill=tk.BOTH)

left_frame = tk.Frame(content_frame, bg=style["frame_bg"])
left_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)

workshop_items_label = tk.Label(left_frame, text="Workshop Items:", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
workshop_items_label.pack(pady=5)

workshop_items_text = scrolledtext.ScrolledText(left_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg=style["highlight_bg"], fg=style["text_fg"])
workshop_items_text.pack(padx=5, pady=5, expand=True, fill=tk.BOTH)

mods_label = tk.Label(left_frame, text="Mods:", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
mods_label.pack(pady=5)

mods_text = scrolledtext.ScrolledText(left_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg=style["highlight_bg"], fg=style["text_fg"])
mods_text.pack(padx=5, pady=5, expand=True, fill=tk.BOTH)

right_frame = tk.Frame(content_frame, bg=style["frame_bg"])
right_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)

urls_label = tk.Label(right_frame, text="Mod Names and Workshop Items:", font=style["label_font"], bg=style["frame_bg"], fg=style["label_fg"])
urls_label.pack(pady=5)

urls_text = scrolledtext.ScrolledText(right_frame, height=20, wrap=tk.WORD, state=tk.DISABLED, bg=style["highlight_bg"], fg=style["text_fg"])
urls_text.pack(padx=5, pady=5, expand=True, fill=tk.BOTH)

button_frame = tk.Frame(frame, bg=style["frame_bg"])
button_frame.pack(pady=10, fill=tk.X)

save_button = tk.Button(button_frame, text="Save to Files", command=save_to_files, bg=style["button_bg"], activebackground=style["button_active_bg"], font=style["font"], fg=style["button_fg"])
save_button.pack(pady=5, fill=tk.X)

clear_button = tk.Button(button_frame, text="Clear", command=clear_text, bg=style["button_bg"], activebackground=style["button_active_bg"], font=style["font"], fg=style["button_fg"])
clear_button.pack(pady=5, fill=tk.X)

exit_button = tk.Button(button_frame, text="Exit", command=exit_app, bg=style["button_bg"], activebackground=style["button_active_bg"], font=style["font"], fg=style["button_fg"])
exit_button.pack(pady=5, fill=tk.X)

# Check for Project Zomboid workshop content on startup
workshop_directory, zomboid_directory = find_project_zomboid_workshop()
if workshop_directory:
    process_directory(workshop_directory)
    steam_dir_label.config(text=f"Steam Installation Directory: {get_steam_install_path()}")
    pzm_dir_label.config(text=f"Project Zomboid Directory: {zomboid_directory}")
else:
    select_directory()

root.mainloop()