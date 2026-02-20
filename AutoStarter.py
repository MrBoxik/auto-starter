import os
import sys
import json
import threading
import subprocess
import time
import tempfile
import locale
import ctypes
from pathlib import Path
from tkinter import Tk, Listbox, Button, Label, filedialog, messagebox, END, SINGLE, Checkbutton, IntVar, Frame, Scrollbar, RIGHT, Y, LEFT, BOTH, PhotoImage

# Try to import tkinterdnd2 (optional) for drag-and-drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

# Try to import win32com to create .lnk shortcuts in Startup
HAVE_PYWIN32 = False
try:
    import pythoncom
    from win32com.shell import shell, shellcon
    HAVE_PYWIN32 = True
except Exception:
    HAVE_PYWIN32 = False

APP_NAME = 'AutoStarter'
APP_USER_MODEL_ID = 'MrBoxik.AutoStarter'
CONFIG_FILENAME = 'items.json'
ICON_FILENAME = 'app_icon.ico'
AUTO_CLOSE_MS = 10_000  # 10 seconds
STARTUP_TASK_NAME = f'{APP_NAME}_Logon'


def get_appdata_dir():
    appdata = os.getenv('APPDATA')
    if not appdata:
        # fallback to user home
        return os.path.join(str(Path.home()), f'.{APP_NAME}')
    return os.path.join(appdata, APP_NAME)


def get_config_path():
    d = get_appdata_dir()
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, CONFIG_FILENAME)


def get_self_path():
    """
    Return the absolute path of the running script or executable.
    Works for normal python and frozen apps (PyInstaller).
    """
    if getattr(sys, 'frozen', False):
        return os.path.abspath(sys.executable)
    return os.path.abspath(sys.argv[0])


def get_resource_base_dir():
    """
    Return location of bundled resources for both script and PyInstaller builds.
    """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return getattr(sys, '_MEIPASS')
    return os.path.dirname(os.path.abspath(__file__))


def find_resource_path(filename):
    """
    Resolve a resource path from common locations.
    """
    if not filename:
        return None
    candidates = [
        os.path.join(get_resource_base_dir(), filename),
        os.path.join(os.path.dirname(get_self_path()), filename),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None


def set_windows_app_user_model_id():
    """
    Ensures the taskbar and title-bar icon are associated with this app on Windows.
    """
    if not sys.platform.startswith('win'):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def get_windows_subprocess_encoding():
    if not sys.platform.startswith('win'):
        return locale.getpreferredencoding(False)
    try:
        cp = ctypes.windll.kernel32.GetOEMCP()
        if cp:
            return f'cp{cp}'
    except Exception:
        pass
    return locale.getpreferredencoding(False)


def run_windows_command(cmd):
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding=get_windows_subprocess_encoding(),
        errors='replace',
    )


def normalize_path(p):
    if not p:
        return p
    # Remove surrounding quotes if any, make absolute
    p = p.strip().strip('"').strip("'")
    try:
        return os.path.abspath(p)
    except Exception:
        return p


def is_self_path(p):
    """
    Return True if p appears to point to this application executable/script.
    We check exact absolute path and also compare basenames (helps with .lnk wrappers).
    """
    try:
        p_norm = normalize_path(p)
        selfp = get_self_path()
        if not p_norm:
            return False
        # exact match
        if os.path.normcase(p_norm) == os.path.normcase(selfp):
            return True
        # same filename (basename) - cautious check
        if os.path.normcase(os.path.basename(p_norm)) == os.path.normcase(os.path.basename(selfp)):
            return True
    except Exception:
        pass
    return False


def load_items():
    """
    Load items from JSON config. If parse fails, return [] and show a warning to user.
    """
    path = get_config_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Basic validation: expect list of dicts or list of strings
        if isinstance(data, list):
            cleaned = []
            for it in data:
                if isinstance(it, dict):
                    p = it.get('path') or ''
                    cleaned.append({'path': normalize_path(p), 'name': it.get('name')})
                else:
                    cleaned.append({'path': normalize_path(str(it))})
            return cleaned
        else:
            return []
    except json.JSONDecodeError as e:
        messagebox.showwarning(APP_NAME, f'Config file is corrupted and could not be read: {e}\nStarting with an empty list.')
        return []
    except Exception as e:
        # unknown error - return empty
        print('load_items error:', e)
        return []


def save_items(items, retries=3):
    """
    Save items using atomic write: write to temp file and then os.replace.
    Returns (True, None) on success or (False, error_message) on failure.
    """
    path = get_config_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_fd = None
    tmp_path = None
    try:
        # Filter/normalize items before saving
        to_save = []
        for it in items:
            if isinstance(it, dict):
                p = normalize_path(it.get('path', ''))
                name = it.get('name') if it.get('name') else None
            else:
                p = normalize_path(str(it))
                name = None
            # skip empty paths and skip self path
            if not p:
                continue
            if is_self_path(p):
                # don't store a path pointing to this app
                continue
            to_save.append({'path': p, **({'name': name} if name else {})})

        # use tempfile in same directory (important for atomic replace on same fs)
        dirpath = os.path.dirname(path) or '.'
        fd, tmp_path = tempfile.mkstemp(prefix=CONFIG_FILENAME, suffix='.tmp', dir=dirpath)
        tmp_fd = fd
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(to_save, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        # atomic replace
        os.replace(tmp_path, path)
        return True, None
    except Exception as e:
        # cleanup
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        return False, str(e)


def open_path(path):
    # Use os.startfile on Windows - works with .lnk, folders, files, programs
    try:
        if sys.platform.startswith('win'):
            os.startfile(path)
        else:
            subprocess.Popen(['xdg-open', path])
    except Exception:
        # fallback: try to run directly
        try:
            subprocess.Popen([path])
        except Exception as e:
            print(f'Failed to open {path}: {e}')


# Startup handling
def get_startup_launch_target_and_args():
    """
    Build the command used by Windows auto-start.
    Returns: (target_executable, args_list)
    """
    if getattr(sys, 'frozen', False):
        return os.path.abspath(sys.executable), ['--nobox']

    script_path = os.path.abspath(sys.argv[0])
    py_exe = os.path.abspath(sys.executable)
    if sys.platform.startswith('win'):
        # Prefer pythonw.exe to avoid opening a console window at logon.
        pyw_exe = os.path.join(os.path.dirname(py_exe), 'pythonw.exe')
        if os.path.basename(py_exe).lower() == 'python.exe' and os.path.exists(pyw_exe):
            py_exe = pyw_exe
    return py_exe, [script_path, '--nobox']


def get_startup_folder():
    # Windows startup folder per-user
    if sys.platform.startswith('win'):
        return os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    else:
        return None


def create_startup_shortcut(target_executable, args_list=None):
    """Try to create a .lnk in the Startup folder pointing to target_executable.
    If pywin32 isn't available, create a .bat wrapper as fallback.
    Return True on success, False on failure.
    """
    if args_list is None:
        args_list = []
    args = subprocess.list2cmdline(args_list)
    startup = get_startup_folder()
    if not startup:
        return False
    name = f'{APP_NAME}.lnk'
    lnk_path = os.path.join(startup, name)
    try:
        if HAVE_PYWIN32:
            # Create a proper .lnk
            shell_link = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
            )
            shell_link.SetPath(target_executable)
            if args:
                shell_link.SetArguments(args)
            # Prefer the executable icon for frozen builds, bundled icon file otherwise.
            if getattr(sys, 'frozen', False):
                icon_path = target_executable
            else:
                icon_path = find_resource_path(ICON_FILENAME)
            if icon_path and os.path.exists(icon_path):
                try:
                    shell_link.SetIconLocation(icon_path, 0)
                except Exception:
                    pass
            persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
            persist_file.Save(lnk_path, 0)
            return True
        else:
            # Fallback: create a .bat that runs the executable
            bat_path = os.path.join(startup, f'{APP_NAME}.bat')
            command = subprocess.list2cmdline([target_executable] + args_list)
            with open(bat_path, 'w', encoding='utf-8') as f:
                # Use start "" "path" to start without bringing up cmd window
                f.write(f'start "" {command}\n')
            return True
    except Exception:
        return False


def create_startup_task(target_executable, args_list=None, quiet=False):
    """
    Prefer Scheduled Task for startup because Startup-folder entries can be delayed by Windows.
    """
    if not sys.platform.startswith('win'):
        return False
    if args_list is None:
        args_list = []

    command = subprocess.list2cmdline([target_executable] + args_list)
    cmd = [
        'schtasks', '/Create',
        '/TN', STARTUP_TASK_NAME,
        '/SC', 'ONLOGON',
        '/TR', command,
        '/F',
        '/RL', 'LIMITED'
    ]
    try:
        p = run_windows_command(cmd)
        if p.returncode == 0:
            return True
        if not quiet:
            msg = ' '.join(part for part in (p.stdout.strip(), p.stderr.strip()) if part).strip()
            if msg:
                print('Failed to create startup task:', msg)
        return False
    except Exception:
        if not quiet:
            print('Failed to create startup task.')
        return False


def startup_task_exists():
    if not sys.platform.startswith('win'):
        return False
    try:
        p = run_windows_command(['schtasks', '/Query', '/TN', STARTUP_TASK_NAME])
        return p.returncode == 0
    except Exception:
        return False


def remove_startup_task():
    if not sys.platform.startswith('win'):
        return False
    if not startup_task_exists():
        return True
    try:
        p = run_windows_command(['schtasks', '/Delete', '/TN', STARTUP_TASK_NAME, '/F'])
        return p.returncode == 0
    except Exception:
        return False


def remove_startup_shortcut():
    startup = get_startup_folder()
    if not startup:
        return False
    lnk = os.path.join(startup, f'{APP_NAME}.lnk')
    bat = os.path.join(startup, f'{APP_NAME}.bat')
    ok = True
    for p in (lnk, bat):
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception as e:
            print('Failed to remove', p, e)
            ok = False
    return ok


def startup_shortcut_exists():
    startup = get_startup_folder()
    if not startup:
        return False
    lnk = os.path.join(startup, f'{APP_NAME}.lnk')
    bat = os.path.join(startup, f'{APP_NAME}.bat')
    return os.path.exists(lnk) or os.path.exists(bat)


def startup_entry_exists():
    return startup_task_exists() or startup_shortcut_exists()


# GUI
class StarterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self._icon_image = None
        self._apply_window_icon()

        self.items = load_items()
        self.auto_close_after_id = None
        self.auto_close_enabled = True
        self.clicked = False

        # Main frame
        frame = Frame(root)
        frame.pack(fill=BOTH, expand=True, padx=6, pady=6)

        # Listbox with scrollbar
        scrollbar = Scrollbar(frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.listbox = Listbox(frame, selectmode=SINGLE)
        self.listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)

        # Buttons
        btn_frame = Frame(root)
        btn_frame.pack(fill='x', padx=6, pady=(0,6))

        Button(btn_frame, text='Add', command=self.add_items).pack(side='left')
        Button(btn_frame, text='Remove', command=self.remove_selected).pack(side='left')
        Button(btn_frame, text='Move Up', command=lambda: self.move_selected(-1)).pack(side='left')
        Button(btn_frame, text='Move Down', command=lambda: self.move_selected(1)).pack(side='left')
        Button(btn_frame, text='Run now', command=self.run_now).pack(side='left')
        Button(btn_frame, text='Save', command=self.save).pack(side='left')

        # Startup checkbox
        self.startup_var = IntVar(value=1 if startup_entry_exists() else 0)
        self.startup_cb = Checkbutton(btn_frame, text='Start with Windows', variable=self.startup_var, command=self.toggle_startup)
        self.startup_cb.pack(side='right')
        self.upgrade_startup_entry_if_needed()

        # Info label
        info = 'Auto Starting of apps, folders, shortcuts and all                     by: MrBoxik'
        Label(root, text=info).pack(fill='x')

        # Fill listbox
        self.refresh_listbox()

        # Bind events
        self.root.bind_all('<Button>', self.on_any_click)
        self.listbox.bind('<Double-1>', self.open_selected)

        # Setup drag-and-drop using tkinterdnd2 when available.
        self.dnd_enabled = False
        if DND_AVAILABLE:
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.on_drop)
                self.listbox.drop_target_register(DND_FILES)
                self.listbox.dnd_bind('<<Drop>>', self.on_drop)
                self.dnd_enabled = True
            except Exception:
                self.dnd_enabled = False
        if not self.dnd_enabled:
            Label(root, text='Drag & drop unavailable: install tkinterdnd2 (pip install tkinterdnd2)').pack(fill='x')

        # Start auto-close timer
        self.start_auto_close_timer()

    def _apply_window_icon(self):
        try:
            icon_path = find_resource_path(ICON_FILENAME)
            if icon_path and sys.platform.startswith('win'):
                self.root.iconbitmap(default=icon_path)
                return

            png_icon = find_resource_path('app_icon.png')
            if png_icon:
                self._icon_image = PhotoImage(file=png_icon)
                self.root.iconphoto(False, self._icon_image)
        except Exception:
            pass

    def refresh_listbox(self):
        self.listbox.delete(0, END)
        for it in self.items:
            p = it.get('path') if isinstance(it, dict) else str(it)
            name = it.get('name') if isinstance(it, dict) and it.get('name') else os.path.basename(p)
            display = name + '    [' + p + ']'
            self.listbox.insert(END, display)

    def add_items(self):
        paths = filedialog.askopenfilenames(title='Select files or shortcuts to add')
        if not paths:
            return
        self._add_paths_to_list(paths)

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.items.pop(idx)
        self.refresh_listbox()

    def move_selected(self, delta):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(self.items):
            return
        self.items[idx], self.items[new_idx] = self.items[new_idx], self.items[idx]
        self.refresh_listbox()
        self.listbox.selection_set(new_idx)

    def open_selected(self, event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        path = self.items[idx].get('path') if isinstance(self.items[idx], dict) else self.items[idx]
        open_path(path)

    def on_drop(self, event):
        data = event.data
        paths = self._parse_dnd_paths(data)
        self._add_paths_to_list(paths)

    def _add_paths_to_list(self, paths):
        added = 0
        for p in paths:
            p_norm = normalize_path(p)
            if is_self_path(p_norm):
                messagebox.showwarning(APP_NAME, "You cannot add the AutoStarter program to its own startup list. Skipping.")
                continue
            self.items.append({'path': p_norm})
            added += 1
        if added:
            self.refresh_listbox()

    def _parse_dnd_paths(self, data):
        # crude parser: split on space unless inside braces
        parts = []
        cur = ''
        in_brace = False
        for ch in data:
            if ch == '{':
                in_brace = True
                cur = ''
            elif ch == '}':
                in_brace = False
                parts.append(cur)
                cur = ''
            elif ch == ' ' and not in_brace:
                if cur:
                    parts.append(cur)
                    cur = ''
            else:
                cur += ch
        if cur:
            parts.append(cur)
        return parts

    def run_now(self):
        threading.Thread(target=self.launch_all).start()

    def launch_all(self):
        """
        Launch all items, skipping anything that appears to be this app itself.
        """
        for it in self.items:
            p = it.get('path') if isinstance(it, dict) else it
            p_norm = normalize_path(p)
            if is_self_path(p_norm):
                print('Skipping self-launch for', p_norm)
                continue
            try:
                open_path(p_norm)
                time.sleep(0.05)
            except Exception as e:
                print('Error launching', p_norm, e)

    def save(self):
        ok, err = save_items(self.items)
        if ok:
            messagebox.showinfo('Saved', 'Saved ' + str(len(self.items)) + ' items.')
        else:
            messagebox.showerror('Error', f'Could not save config: {err}')

    def toggle_startup(self):
        if self.startup_var.get():
            target, args_list = get_startup_launch_target_and_args()
            # Prevent creating startup entry that points to the app being launched by the app list
            if any(is_self_path(it.get('path')) for it in self.items if isinstance(it, dict)):
                messagebox.showwarning(APP_NAME, "Your list contains a path to this AutoStarter program. Enabling 'Start with Windows' while your list contains the AutoStarter itself could cause loops. Please remove it from the list first.")
                self.startup_var.set(0)
                return
            ok_task = create_startup_task(target, args_list, quiet=True)
            if ok_task:
                # Clean up old startup shortcut if we successfully moved to Scheduled Task.
                remove_startup_shortcut()
                messagebox.showinfo('Startup', 'Enabled Start with Windows using Task Scheduler (runs immediately at logon).')
                return

            ok_shortcut = create_startup_shortcut(target, args_list)
            if not ok_shortcut:
                messagebox.showwarning('Startup', 'Could not create startup shortcut. You may need to run as admin or check permissions.')
                self.startup_var.set(0)
            else:
                messagebox.showinfo('Startup', 'Enabled Start with Windows using Startup folder (fallback mode).')
        else:
            ok_task = remove_startup_task()
            ok_shortcut = remove_startup_shortcut()
            if ok_task and ok_shortcut:
                messagebox.showinfo('Startup', 'Disabled Start with Windows.')
            else:
                messagebox.showwarning('Startup', 'Could not fully remove startup entries (task and/or shortcut).')

    def upgrade_startup_entry_if_needed(self):
        """
        Upgrade legacy Startup-folder auto-start to Scheduled Task silently.
        """
        if not self.startup_var.get():
            return
        # Keep existing Startup-folder entries untouched to avoid repeated
        # permission errors on systems that do not allow task creation.
        if startup_shortcut_exists():
            return
        if startup_task_exists():
            return

        target, args_list = get_startup_launch_target_and_args()
        if create_startup_task(target, args_list, quiet=True):
            remove_startup_shortcut()

    def start_auto_close_timer(self):
        # schedule the auto-close
        if self.auto_close_enabled:
            self.auto_close_after_id = self.root.after(AUTO_CLOSE_MS, self.auto_launch_and_exit)

    def cancel_auto_close(self):
        if self.auto_close_after_id:
            try:
                self.root.after_cancel(self.auto_close_after_id)
            except Exception:
                pass
            self.auto_close_after_id = None

    def on_any_click(self, event=None):
        # When the user clicks anywhere inside the app, cancel auto-close so they can edit
        if not self.clicked:
            self.clicked = True
            self.cancel_auto_close()

    def auto_launch_and_exit(self):
        # Called when 10s passed with no click
        # Launch everything and close app
        threading.Thread(target=self._launch_and_exit_worker).start()

    def _launch_and_exit_worker(self):
        self.launch_all()
        # wait a short moment then exit GUI
        time.sleep(0.2)
        try:
            self.root.quit()
        except Exception:
            pass


def main():
    set_windows_app_user_model_id()
    # If DND_AVAILABLE and tkinterdnd2 exists, create TkinterDnD.Tk, otherwise normal Tk
    if DND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = Tk()
    else:
        root = Tk()

    app = StarterApp(root)
    # If the program was launched with --run (or from startup), we still show GUI for 10s and then auto-launch
    # But user can also pass --nobox to skip GUI (headless launch) if they like
    if '--nobox' in sys.argv:
        # headless: launch and exit immediately
        app.launch_all()
        return

    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    finally:
        # Save before exit
        save_items(app.items)


if __name__ == '__main__':
    main()
