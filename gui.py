#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tkinter GUI for the Email Archive Tool.

Reuses every function from archive.py - the command-line tool is left
untouched and stays fully functional. Launch with:  python gui.py
"""

import argparse
import csv
import os
import queue
import re
import shutil
import sys
import threading
import traceback
import webbrowser

import archive

RING = []
LOG_TO_UI = []     # functions called with each log line (set by the GUI)
LOG_LIMIT = 2000
AZURE_APPS_URL = 'https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/Applications'
OUTLOOK_IMAP_URL = 'https://outlook.live.com/mail/0/options/mail/forwarding'
GMAIL_APP_PASSWORD_URL = 'https://support.google.com/mail/answer/185833'
GMAIL_IMAP_URL = 'https://support.google.com/mail/answer/7126229'
YAHOO_APP_PASSWORD_URL = 'https://help.yahoo.com/kb/generate-third-party-passwords-sln15241.html'
APP_NAME = 'Email Archive Tool'
MANUAL_FILE = 'EmailArchiveTool_Manual.pdf'
ARCHIVE_COLUMNS = {
    'ref': 'Ref No.', 'date': 'Transaction Date', 'agency': 'Agency',
    'sender': 'Sender Name', 'subject': 'Subject', 'action': 'Action',
    'response': 'Response Date', 'attachments': 'Attachments',
}
ARCHIVE_HEADINGS = {
    'ref': 'Ref No.', 'date': 'Date', 'agency': 'Agency', 'sender': 'Sender',
    'subject': 'Subject', 'action': 'Action', 'response': 'Response',
    'attachments': 'Attachments',
}
LIGHT_THEME = {
    'bg': '#eef1f6', 'panel': '#f7f9fc', 'field': '#ffffff', 'border': '#d7dde6',
    'text': '#1f2933', 'muted': '#5b6672', 'accent': '#0f6cbd', 'accent_dark': '#0a5493',
    'danger': '#b3261e', 'thread_bg': '#e6edf7', 'alt': '#f7f9fc', 'selected': '#d6e6f7',
    'link': '#1a0dab', 'heading': '#12395e',
    'hover': '#e6ebf3', 'pressed': '#dbe2ec', 'border_active': '#c8d2e0',
    'disabled_fg': '#c9d7e6',
}
DARK_THEME = {
    'bg': '#121417', 'panel': '#1b1f24', 'field': '#23272d', 'border': '#333a42',
    'text': '#e6e9ee', 'muted': '#98a2b3', 'accent': '#4d9de0', 'accent_dark': '#3b84c4',
    'danger': '#e0807a', 'thread_bg': '#1d2836', 'alt': '#1a1e24', 'selected': '#2b3c50',
    'link': '#6cb4ee', 'heading': '#9cc4ee',
    'hover': '#262b32', 'pressed': '#2e343c', 'border_active': '#3d454f',
    'disabled_fg': '#5b6672',
}
THEME = dict(LIGHT_THEME)


def set_palette(dark):
    THEME.clear()
    THEME.update(DARK_THEME if dark else LIGHT_THEME)
GUI_PROVIDERS = {
    'outlook': {'label': 'Outlook.com / Office 365', 'server': 'outlook.office365.com',
                'port': 993, 'ssl': True},
    'gmail': {'label': 'Gmail', 'server': 'imap.gmail.com', 'port': 993, 'ssl': True},
    'yahoo': {'label': 'Yahoo Mail', 'server': 'imap.mail.yahoo.com', 'port': 993, 'ssl': True},
    'custom': {'label': 'Other IMAP server', 'server': '', 'port': 993, 'ssl': True},
}


def provider_for_server(server):
    server = str(server or '').strip().lower()
    for key, preset in GUI_PROVIDERS.items():
        if preset.get('server') and server == preset['server'].lower():
            return key
    return 'custom'


def apply_theme(win, dark=False):
    import tkinter as tk
    from tkinter import font as tkfont
    from tkinter import ttk
    set_palette(dark)
    c = THEME
    families = set(tkfont.families(win))
    ui = 'Segoe UI' if 'Segoe UI' in families else 'TkDefaultFont'
    win.configure(bg=c['bg'])
    style = ttk.Style(win)
    try:
        style.theme_use('clam')
    except tk.TclError:
        pass
    style.configure('.', font=(ui, 10), background=c['bg'], foreground=c['text'])
    style.configure('TFrame', background=c['bg'])
    style.configure('TLabel', background=c['bg'], foreground=c['text'])
    style.configure('TLabelframe', background=c['bg'], bordercolor=c['border'],
                    relief='flat', lightcolor=c['border'], darkcolor=c['border'])
    style.configure('TLabelframe.Label', background=c['bg'], foreground=c['accent'],
                    font=(ui, 10, 'bold'))
    style.configure('TButton', background=c['panel'], foreground=c['text'], padding=(14, 7),
                    bordercolor=c['border'], lightcolor=c['border'], darkcolor=c['border'],
                    focuscolor=c['accent'])
    style.map('TButton',
              background=[('active', c['hover']), ('pressed', c['pressed'])],
              bordercolor=[('active', c['border_active'])])
    style.configure('Accent.TButton', background=c['accent'], foreground='#ffffff',
                    padding=(14, 7), bordercolor=c['accent'],
                    lightcolor=c['accent'], darkcolor=c['accent'])
    style.map('Accent.TButton',
              background=[('active', c['accent_dark']), ('pressed', c['accent_dark'])],
              foreground=[('disabled', c['disabled_fg'])])
    for kind in ('TEntry', 'TSpinbox', 'TCombobox'):
        style.configure(kind, fieldbackground=c['field'], background=c['field'],
                        foreground=c['text'], insertcolor=c['text'], bordercolor=c['border'],
                        lightcolor=c['border'], darkcolor=c['border'], padding=(8, 6))
    style.map('TCombobox',
              fieldbackground=[('readonly', c['field'])],
              foreground=[('readonly', c['text'])],
              selectbackground=[('readonly', c['field'])],
              selectforeground=[('readonly', c['text'])])
    style.configure('TCheckbutton', background=c['bg'], foreground=c['text'])
    style.map('TCheckbutton', background=[('active', c['bg'])])
    style.configure('TNotebook', background=c['border'], borderwidth=0, tabmargins=(2, 4, 2, 0))
    style.configure('TNotebook.Tab', background=c['panel'], foreground=c['muted'],
                    padding=(14, 7), borderwidth=0)
    style.map('TNotebook.Tab',
              background=[('selected', c['field'])],
              foreground=[('selected', c['text'])])
    style.configure('Vertical.TScrollbar', background=c['panel'], troughcolor=c['bg'],
                    bordercolor=c['bg'], arrowsize=12, borderwidth=0)
    style.configure('Horizontal.TScrollbar', background=c['panel'], troughcolor=c['bg'],
                    bordercolor=c['bg'], arrowsize=12, borderwidth=0)
    style.configure('Treeview', background=c['field'], fieldbackground=c['field'],
                    foreground=c['text'], bordercolor=c['border'], borderwidth=0,
                    rowheight=28, font=(ui, 10))
    style.configure('Treeview.Heading', background=c['panel'], foreground=c['accent'],
                    font=(ui, 10, 'bold'), borderwidth=0)
    style.map('Treeview.Heading', background=[('active', c['hover'])])
    style.map('Treeview',
              background=[('selected', c['selected'])],
              foreground=[('selected', c['text'])])
    return ui


def install_log_hook():
    """Route archive.say/warn/fail into the GUI log (and the console)."""
    global LOG_TO_UI

    def emit(message=''):
        text = str(message)
        RING.append(text)
        del RING[:-LOG_LIMIT]
        try:
            print(text)
        except Exception:
            pass
        for sink in list(LOG_TO_UI):
            try:
                sink(text)
            except Exception:
                LOG_TO_UI.remove(sink)

    archive.say = emit
    archive.warn = lambda message: emit('[warn] ' + str(message))
    archive.fail = lambda message: emit('[error] ' + str(message))


def reconcile_keywords(existing, items, hide_keys):
    """exclude_keywords list after applying the GUI's Hide checkboxes.

    hide_keys: set of sender keywords currently checked in the UI.
    Keywords tied to a visible email (its name and its address) are added
    when checked and removed when unchecked; manually-added keywords stay
    unless a visible email is unchecked.
    """
    words = list(existing or [])
    for item in items:
        keywords_list = archive.keywords_for(item)
        if not keywords_list:
            continue
        if any(keyword in hide_keys for keyword in keywords_list):
            for keyword in keywords_list:
                if keyword not in words:
                    words.append(keyword)
        else:
            words = [word for word in words if word not in keywords_list]
    return words


def selected_positions(summaries, selected_keys):
    return [index + 1 for index, item in enumerate(summaries)
            if item.get('key') in selected_keys]


def _log_rows(csv_path):
    if not os.path.exists(csv_path):
        return []
    try:
        with open(csv_path, newline='', encoding='utf-8-sig') as handle:
            return list(csv.DictReader(handle))
    except OSError:
        return []


def attachment_names(value):
    text = str(value or '').strip()
    if not text or text == archive.DASH:
        return []
    return [part.strip() for part in text.split(';') if part.strip()]


def _ref_files(ref, root, rows, match, include_eml):
    files = []
    used = set()

    def add(label, kind, path):
        if path in used:
            return
        used.add(path)
        files.append((label, kind, path))

    for row in rows:
        if not match(row):
            continue
        for name in attachment_names(row.get('Attachments')):
            path = os.path.join(root, 'Attachments', name)
            if os.path.exists(path):
                add(name, 'attachment', path)
    eml_dir = os.path.join(root, 'EML')
    if include_eml and os.path.isdir(eml_dir):
        for name in sorted(os.listdir(eml_dir)):
            if name.endswith('.eml') and (name == ref + '.eml' or name.startswith(ref + '-')):
                add(name, 'EML', os.path.join(eml_dir, name))
    pdf = os.path.join(root, 'PDF', ref + '.pdf')
    if os.path.exists(pdf):
        add(ref + '.pdf', 'PDF', pdf)
    return files


def files_for_message(ref, sender_email, date, subject, root):
    rows = _log_rows(os.path.join(root, 'log.csv'))
    return _ref_files(ref, root, rows, lambda row: (
        row.get('Ref No.') == ref
        and row.get('Sender Email') == sender_email
        and row.get('Transaction Date') == date
        and row.get('Subject') == subject), False)


def files_for_thread(ref, root):
    rows = _log_rows(os.path.join(root, 'log.csv'))
    return _ref_files(ref, root, rows, lambda row: row.get('Ref No.') == ref, True)


def ui_rows(summaries, selected_keys, hidden_names, thread_view, filter_text='',
            date_from='', date_to='', unread_only=False, has_attachment=False,
            archived=None):
    """Flattened rows for the Tkinter treeview.

    Returns a list of dicts with 'kind' = 'thread' (header) or 'mail'.
    Selection/hide flags are resolved here so the tree code stays thin.
    Non-matching messages are hidden (thread headers are shown whenever at
    least one member matches).
    """
    lower = (filter_text or '').strip().lower()
    d_from = (date_from or '').strip()
    d_to = (date_to or '').strip()

    def matches(item):
        if lower:
            sender = (item.get('from_name') or item.get('from_email') or '').lower()
            subject = (item.get('subject') or '').lower()
            if lower not in sender and lower not in subject:
                return False
        date_value = str(item.get('date') or '').strip()
        if d_from and date_value and date_value < d_from:
            return False
        if d_to and date_value and date_value > d_to:
            return False
        if unread_only and item.get('seen', True):
            return False
        if has_attachment and not item.get('has_attachment'):
            return False
        if archived is not None and bool(item.get('archived')) != archived:
            return False
        return True

    rows = []
    shown_threads = set()
    for index, item in enumerate(summaries):
        thread_no = item.get('thread_no') or 0
        if not matches(item):
            continue
        if thread_view and thread_no and thread_no not in shown_threads:
            shown_threads.add(thread_no)
            member_count = sum(1 for other in summaries if (other.get('thread_no') or 0) == thread_no)
            rows.append({
                'kind': 'thread', 'iid': 'thread-%d' % thread_no, 'thread_no': thread_no,
                'count': member_count,
            })
        hidden = any(hide in hidden_names for hide in archive.keywords_for(item))
        rows.append({
            'kind': 'mail', 'iid': 'mail-%d' % index, 'index': index, 'item': item,
            'thread_no': thread_no,
            'date': item.get('date') or '-',
            'sender': item.get('from_name') or item.get('from_email') or '(unknown sender)',
            'subject': item.get('subject') or '(no subject)',
            'selected': item['key'] in selected_keys,
            'hidden': hidden,
            'unread': not bool(item.get('seen', True)),
            'archived': bool(item.get('archived')),
        })
    return rows


class SettingsDialog:
    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk
        self.app = app
        self.window = tk.Toplevel(app.root_window)
        self.window.title('Settings - email, Azure and IMAP')
        self.window.transient(app.root_window)
        self.window.resizable(True, True)
        self.window.configure(bg=THEME['bg'])
        self.links = []
        config = app.config

        buttons = ttk.Frame(self.window, padding=(12, 0, 12, 12))
        buttons.pack(side='bottom', fill='x')
        ttk.Button(buttons, text='Save', style='Accent.TButton',
                   command=self.on_save).pack(side='right', padx=(6, 0))
        ttk.Button(buttons, text='Cancel', command=self.window.destroy).pack(side='right')

        outer = ttk.Frame(self.window)
        outer.pack(side='top', fill='both', expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0, bg=THEME['bg'])
        vbar = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        body = ttk.Frame(canvas, padding=12)
        body.columnconfigure(1, weight=1)
        window_item = canvas.create_window((0, 0), window=body, anchor='nw')

        def _on_body_config(event):
            canvas.configure(scrollregion=canvas.bbox('all'))

        def _on_canvas_config(event):
            canvas.itemconfigure(window_item, width=event.width)

        def _on_wheel(event):
            canvas.yview_scroll(int(-event.delta / 120), 'units')

        body.bind('<Configure>', _on_body_config)
        canvas.bind('<Configure>', _on_canvas_config)

        def add_field(row, label, widget):
            ttk.Label(body, text=label).grid(row=row, column=0, sticky='e', padx=(0, 8), pady=3)
            widget.grid(row=row, column=1, sticky='we', pady=3)

        provider = config.get('imap_provider')
        if provider not in GUI_PROVIDERS:
            provider = provider_for_server(config.get('imap_server'))
        self.var_provider = _var(GUI_PROVIDERS[provider]['label'])
        provider_box = ttk.Combobox(body, textvariable=self.var_provider, state='readonly', width=38,
                                    values=[preset['label'] for preset in GUI_PROVIDERS.values()])
        add_field(0, 'Mail provider:', provider_box)
        provider_box.bind('<<ComboboxSelected>>', lambda event: self.on_provider_changed())

        self.var_email = _var(config.get('email') or '')
        add_field(1, 'Email address:', ttk.Entry(body, textvariable=self.var_email, width=40))

        self.var_app_password = _var(config.get('app_password') or '')
        self.app_pw_entry = ttk.Entry(body, textvariable=self.var_app_password, width=40, show='*')
        add_field(2, 'App password:', self.app_pw_entry)

        self.var_server = _var(config.get('imap_server') or 'outlook.office365.com')
        add_field(3, 'IMAP server:', ttk.Entry(body, textvariable=self.var_server, width=40))

        self.var_port = _var(int(config.get('imap_port') or 993))
        add_field(4, 'IMAP port:', ttk.Spinbox(body, from_=1, to=65535, textvariable=self.var_port, width=38))

        self.var_ssl = tk.IntVar(value=1 if config.get('use_ssl', True) else 0)
        add_field(5, 'Use SSL:', ttk.Checkbutton(body, variable=self.var_ssl, text=' secure connection (on by default)'))

        self.var_client_id = _var(config.get('client_id') or '')
        self.client_id_entry = ttk.Entry(body, textvariable=self.var_client_id, width=40)
        add_field(6, 'Azure app (client) ID:', self.client_id_entry)

        self.var_tenant = _var(str(config.get('tenant') or 'common'))
        self.tenant_box = ttk.Combobox(body, textvariable=self.var_tenant, state='readonly', width=38,
                                       values=('common', 'consumers', 'organizations'))
        add_field(7, 'Tenant:', self.tenant_box)

        self.var_limit = _var(int(config.get('fetch_limit') or 50))
        add_field(8, 'Fetch limit:', ttk.Spinbox(body, from_=1, to=5000, textvariable=self.var_limit, width=38))

        self.var_threads = tk.IntVar(value=1 if config.get('thread_view', True) else 0)
        add_field(9, 'Conversation threads:', ttk.Checkbutton(body, variable=self.var_threads, text=' group replies under one Ref No.'))

        self.lbl_auth = ttk.Label(body, wraplength=430, justify='left', foreground=THEME['heading'])
        self.lbl_auth.grid(row=10, column=0, columnspan=2, sticky='we', pady=(6, 0))

        note = ('Changing the email or the Azure app ID invalidates the saved Microsoft '
                'sign-in - the next fetch will ask you to sign in again, or delete '
                'token_cache.json next to the program to force it now. Passwords are stored '
                'only in config.json next to the program, never sent anywhere except your '
                'own mail server.')
        ttk.Label(body, text=note, wraplength=430, foreground=THEME['muted']).grid(
            row=11, column=0, columnspan=2, sticky='we', pady=(8, 2))

        help_frame = ttk.Frame(body)
        help_frame.grid(row=12, column=0, columnspan=2, sticky='we', pady=(10, 2))
        ttk.Label(help_frame, text='Where to find these values:',
                  font=((self.app.ui_font or 'TkDefaultFont'), 9, 'bold')).pack(anchor='w')

        def help_item(text, link_text, url):
            row = ttk.Frame(help_frame)
            row.pack(fill='x', pady=(4, 0))
            ttk.Label(row, text=text, wraplength=430, justify='left',
                      foreground=THEME['muted']).pack(anchor='w')
            link = _link(row, '\u2794  ' + link_text, url)
            self.links.append((link, url))
            link.pack(anchor='w', padx=(12, 0), pady=(1, 0))

        help_item('For Outlook.com - Azure app (client) ID: Azure portal > App registrations > '
                  'New registration, account type "Personal Microsoft accounts only"; on its '
                  'Authentication page enable "Allow public client flows"; copy the '
                  'Application (client) ID from the Overview page.',
                  'Open the Azure app registrations page', AZURE_APPS_URL)
        help_item('For Outlook.com - enable IMAP in Outlook Settings > Mail > Forwarding and '
                  'IMAP: tick "Let devices and apps use IMAP". If fetching fails with '
                  '"User is authenticated but not connected", turn it OFF, then ON again, '
                  'wait about a minute, and fetch again.',
                  'Open the Outlook IMAP settings page', OUTLOOK_IMAP_URL)
        help_item('For Gmail - IMAP must be on: sign in at Gmail on a browser > Settings > '
                  'See all settings > Forwarding and POP/IMAP > Enable IMAP > Save.',
                  'Gmail instructions to enable IMAP', GMAIL_IMAP_URL)
        help_item('For Gmail - app password: enable 2-step verification on your Google '
                  'account, then create a 16-character app password for Mail and paste it in '
                  'the App password field above. Your normal Gmail password will not work.',
                  'Create a Gmail app password', GMAIL_APP_PASSWORD_URL)
        help_item('For Yahoo - app password: in Yahoo Account Security choose "Generate third '
                  'party app passwords" (app: Mail for Desktop) and paste the generated '
                  'password in the App password field above.',
                  'Create a Yahoo app password', YAHOO_APP_PASSWORD_URL)
        ttk.Label(help_frame, text='For any other IMAP server (Custom): use the address, port '
                  'and credentials your provider gives you - normally an app password, not '
                  'your account login, is required by Gmail, Yahoo and company servers.',
                  wraplength=430, justify='left', foreground=THEME['muted']).pack(anchor='w', pady=(6, 0))

        self.on_provider_changed()

        def _bind_wheel(widget):
            widget.bind('<MouseWheel>', _on_wheel)
            for child in widget.winfo_children():
                _bind_wheel(child)

        _bind_wheel(self.window)
        self.window.update_idletasks()
        content_w = body.winfo_reqwidth() + vbar.winfo_reqwidth() + 30
        content_h = body.winfo_reqheight() + buttons.winfo_reqheight() + 30
        sw = self.window.winfo_screenwidth()
        sh = self.window.winfo_screenheight()
        self.window.geometry(f'{min(content_w, sw - 40)}x{min(content_h, sh - 100)}')
        self.window.minsize(min(420, content_w), min(320, sh - 140))

    def provider_key(self):
        label = self.var_provider.get()
        for key, preset in GUI_PROVIDERS.items():
            if preset['label'] == label:
                return key
        return 'custom'

    def on_provider_changed(self):
        key = self.provider_key()
        preset = GUI_PROVIDERS[key]
        if key != 'custom':
            self.var_server.set(preset['server'])
            self.var_port.set(preset['port'])
            self.var_ssl.set(1 if preset['ssl'] else 0)
        oauth = key == 'outlook'
        auth = not oauth
        self.app_pw_entry.state(['!disabled'] if auth else ['disabled'])
        state = ['!disabled'] if oauth else ['disabled']
        self.client_id_entry.state(state)
        self.tenant_box.state(state)
        if key == 'outlook':
            self.lbl_auth.config(text=(
                'Sign-in method: Microsoft OAuth2, using the Azure app (client) ID below - '
                'no password is needed or stored. On the first fetch a code appears; enter '
                'it at microsoft.com/devicelogin and approve.'))
        elif key == 'gmail':
            self.lbl_auth.config(text=(
                'Sign-in method: your Gmail address + an App password (your normal password '
                'will NOT work). Enable 2-step verification, then create the App password '
                'from the link in "Where to find these values" and paste it above.'))
        elif key == 'yahoo':
            self.lbl_auth.config(text=(
                'Sign-in method: your Yahoo address + a Yahoo app password. Generate it from '
                'the link in "Where to find these values" and paste it above.'))
        else:
            self.lbl_auth.config(text=(
                'Sign-in method: this email + the password above, exactly as your mail '
                'server requires. Gmail and Yahoo need an app password, most others accept '
                'the account password or an application-specific one.'))

    def collect(self):
        return {
            'email': self.var_email.get(),
            'app_password': self.var_app_password.get(),
            'client_id': self.var_client_id.get(),
            'tenant': self.var_tenant.get(),
            'imap_server': self.var_server.get(),
            'imap_port': self.var_port.get(),
            'use_ssl': bool(self.var_ssl.get()),
            'fetch_limit': self.var_limit.get(),
            'thread_view': bool(self.var_threads.get()),
            'imap_provider': self.provider_key(),
        }

    def on_save(self):
        if self.app.apply_settings(self.collect()):
            self.window.destroy()


class MailToolApp:
    TICK = 80

    def __init__(self, root_window, config_path):
        self.root_window = root_window
        self.config_path = os.path.abspath(config_path)
        self.queue = queue.Queue()
        self.config = archive.load_config(self.config_path)
        self.dark = bool(self.config.get('dark_mode', False))
        self.ui_font = apply_theme(root_window, self.dark)
        self.folder = archive.resolve_root(self.config)
        archive.ensure_dirs(self.folder)
        self.summaries = []
        self.selected_keys = set()
        self.hidden_names = set(self.config.get('exclude_keywords') or [])
        self.seen = archive.load_seen(self.folder)
        self.rowmap = {}
        self.log_rows = []
        self.archive_filtered = []
        self.arch_sort_key = ''
        self.arch_sort_desc = False
        self.worker = None
        self.build_widgets()
        LOG_TO_UI.append(self.log)
        self.status(f'Ready - archive folder: {self.folder}')
        self.log('Settings loaded from config.json.')
        self.root_window.protocol('WM_DELETE_WINDOW', self.on_close)
        self.root_window.after(self.TICK, self.poll)

    # ------------------------------------------------------------------ UI
    def _apply_initial_geometry(self):
        win = self.root_window
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        width = min(1360, max(900, sw - 120))
        height = min(1040, max(680, sh - 130))
        width = min(width, sw)
        height = min(height, sh)
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height) // 3)
        win.geometry(f'{width}x{height}+{x}+{y}')
        win.minsize(min(900, sw), min(600, sh))

    def build_widgets(self):
        import tkinter as tk
        from tkinter import ttk
        self.root_window.title(APP_NAME)
        self._apply_initial_geometry()

        settings = ttk.LabelFrame(self.root_window, text=' Email and connection ', padding=8)
        settings.pack(fill='x', padx=10, pady=(8, 4))
        settings.columnconfigure(1, weight=1)
        ttk.Label(settings, text='Email:').grid(row=0, column=0, sticky='e', padx=(2, 6))
        self.var_email = _var(self.config.get('email') or '')
        ttk.Entry(settings, textvariable=self.var_email).grid(row=0, column=1, sticky='we')
        ttk.Button(settings, text='Settings...', width=14,
                   command=self.open_settings).grid(row=0, column=2, padx=(10, 2))
        ttk.Label(settings, text='Azure sign-in, IMAP server and limits are in Settings.',
                  foreground=THEME['muted']).grid(
            row=1, column=0, columnspan=3, sticky='w', pady=(8, 0))

        body = ttk.Frame(self.root_window)
        body.pack(fill='both', expand=True, padx=10, pady=(0, 8))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        content = ttk.Frame(body)
        content.grid(row=0, column=0, sticky='nsew')

        sidebar = ttk.Frame(body, width=210)
        sidebar.grid(row=0, column=1, sticky='ns', padx=(12, 0))
        sidebar.grid_propagate(False)
        ttk.Label(sidebar, text='Help', font=(self.ui_font, 11, 'bold')).pack(
            anchor='w', fill='x', pady=(8, 6))
        self.btn_manual = ttk.Button(sidebar, text='Open PDF Manual', style='Accent.TButton',
                                     command=self.open_manual)
        self.btn_manual.pack(fill='x', pady=4)
        self.btn_theme = ttk.Button(sidebar, text='Use dark mode' if not self.dark else 'Use light mode',
                                    command=self.toggle_theme)
        self.btn_theme.pack(fill='x', pady=4)
        ttk.Button(sidebar, text='About', command=self.about).pack(fill='x', pady=4)
        ttk.Button(sidebar, text='Uninstall', command=self.uninstall).pack(fill='x', pady=4)
        sidebar_note = ('Requires a step-by-step PDF guide with screenshots for every '
                        'screen, saved as EmailArchiveTool_Manual.pdf next to the program. '
                        'Uninstall removes all saved data.')
        ttk.Label(sidebar, text=sidebar_note, wraplength=190, justify='left',
                  foreground=THEME['muted']).pack(fill='x', pady=(12, 0))

        bar = ttk.Frame(content)
        bar.pack(fill='x', pady=(2, 6))
        ttk.Button(bar, text='Fetch emails', width=14, style='Accent.TButton',
                   command=self.fetch).pack(side='left', padx=(0, 8))
        ttk.Button(bar, text='Archive selected', width=16, style='Accent.TButton',
                   command=self.archive).pack(side='left', padx=8)
        for text, command in (('Export Excel', self.export_excel), ('Clear log', self.clear_log)):
            ttk.Button(bar, text=text, width=14, command=command).pack(side='left', padx=8)
        ttk.Button(bar, text='Open folder', width=12,
                   command=lambda: _open_path(self.folder)).pack(side='left', padx=8)
        ttk.Button(bar, text='Open config', width=12,
                   command=lambda: _open_path(self.config_path)).pack(side='left', padx=8)

        self.status_var = _var('')
        ttk.Label(content, textvariable=self.status_var, anchor='w').pack(
            side='bottom', fill='x', pady=(0, 4))

        notes = ttk.Notebook(content)
        notes.pack(fill='both', expand=True, pady=4)

        emails_frame = ttk.Frame(notes)
        notes.add(emails_frame, text='Emails')
        filter_row = ttk.Frame(emails_frame)
        filter_row.grid(row=0, column=0, columnspan=2, sticky='we', pady=(0, 3))
        ttk.Label(filter_row, text='Filter:').pack(side='left')
        self.var_filter = _var('')
        entry = ttk.Entry(filter_row, textvariable=self.var_filter, width=40)
        entry.pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(filter_row, text='Clear all', command=self.clear_email_filters).pack(side='left')
        self.var_filter.trace_add('write', lambda *args: self.refresh_tree())
        extras = ttk.Frame(emails_frame)
        extras.grid(row=1, column=0, columnspan=2, sticky='we', pady=(0, 3))
        ttk.Label(extras, text='From:').pack(side='left')
        self.var_email_from = _var('')
        ttk.Entry(extras, textvariable=self.var_email_from, width=11).pack(side='left', padx=(2, 8))
        ttk.Label(extras, text='To:').pack(side='left')
        self.var_email_to = _var('')
        ttk.Entry(extras, textvariable=self.var_email_to, width=11).pack(side='left', padx=(2, 8))
        self.var_unread_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text='Unread only',
                        variable=self.var_unread_only).pack(side='left', padx=(0, 8))
        self.var_has_attach = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text='Has attachment',
                        variable=self.var_has_attach).pack(side='left', padx=(0, 8))
        ttk.Label(extras, text='Archived:').pack(side='left')
        self.var_archived = _var('Any')
        self.archived_box = ttk.Combobox(extras, textvariable=self.var_archived, width=12,
                                         state='readonly', values=('Any', 'Archived', 'Not archived'))
        self.archived_box.pack(side='left', padx=(2, 0))
        for var in (self.var_email_from, self.var_email_to, self.var_archived,
                    self.var_unread_only, self.var_has_attach):
            var.trace_add('write', lambda *args: self.refresh_tree())
        self.tree = ttk.Treeview(emails_frame, columns=('sel', 'hide', 'date', 'sender', 'subject'),
                                 show='headings', selectmode='none')
        self.tree.heading('sel', text='Keep')
        self.tree.heading('hide', text='Hide')
        self.tree.heading('date', text='Date')
        self.tree.heading('sender', text='Sender')
        self.tree.heading('subject', text='Subject')
        self.tree.column('sel', width=56, anchor='center', stretch=False)
        self.tree.column('hide', width=56, anchor='center', stretch=False)
        self.tree.column('date', width=110, stretch=False)
        self.tree.column('sender', width=230, stretch=False)
        self.tree.column('subject', width=460, stretch=True)
        tree_y = ttk.Scrollbar(emails_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_y.set)
        self.tree.grid(row=2, column=0, sticky='nsew')
        tree_y.grid(row=2, column=1, sticky='ns')
        emails_frame.rowconfigure(2, weight=1)
        emails_frame.columnconfigure(0, weight=1)
        self.tree.tag_configure('thread', background=THEME['thread_bg'], foreground=THEME['heading'],
                                font=(self.ui_font, 9, 'bold'))
        self.tree.tag_configure('solo', font=(self.ui_font, 9, 'bold'))
        self.tree.tag_configure('unread', font=(self.ui_font, 9, 'bold'), foreground=THEME['danger'])
        self.tree.tag_configure('alt', background=THEME['alt'])
        self.tree.bind('<Button-1>', self.on_click)
        kw_frame = ttk.Frame(emails_frame)
        kw_frame.grid(row=3, column=0, columnspan=2, sticky='we', pady=(4, 0))
        ttk.Label(kw_frame, text='Hidden senders (exclude_keywords):').pack(side='left')
        self.keywords = tk.Listbox(kw_frame, height=3, exportselection=False)
        self.keywords.pack(side='left', fill='x', expand=True, padx=6)
        self.keywords.configure(bg=THEME['field'], fg=THEME['text'],
                                selectbackground=THEME['accent'], selectforeground='#ffffff',
                                highlightthickness=0, font=(self.ui_font, 9))
        ttk.Button(kw_frame, text='Remove', command=self.remove_keyword).pack(side='left', padx=2)
        ttk.Button(kw_frame, text='Clear all', command=self.clear_keywords).pack(side='left', padx=2)

        archive_frame = ttk.Frame(notes)
        notes.add(archive_frame, text='Archive')
        arch_bar = ttk.Frame(archive_frame)
        arch_bar.grid(row=0, column=0, columnspan=2, sticky='we', pady=(0, 3))
        ttk.Button(arch_bar, text='Refresh', command=self.refresh_archive).pack(side='left')
        ttk.Button(arch_bar, text='Import CSV', command=self.import_archive_csv).pack(side='left', padx=4)
        self.arch_count = _var('')
        ttk.Label(arch_bar, textvariable=self.arch_count).pack(side='left', padx=8)
        ttk.Label(arch_bar, text='Double-click a row to open its PDF.').pack(side='left', padx=8)
        arch_filters = ttk.Frame(archive_frame)
        arch_filters.grid(row=1, column=0, columnspan=2, sticky='we', pady=(0, 3))
        ttk.Label(arch_filters, text='Search:').pack(side='left')
        self.var_arch_search = _var('')
        ttk.Entry(arch_filters, textvariable=self.var_arch_search, width=20).pack(side='left', padx=(2, 8))
        ttk.Label(arch_filters, text='From:').pack(side='left')
        self.var_arch_from = _var('')
        ttk.Entry(arch_filters, textvariable=self.var_arch_from, width=11).pack(side='left', padx=(2, 8))
        ttk.Label(arch_filters, text='To:').pack(side='left')
        self.var_arch_to = _var('')
        ttk.Entry(arch_filters, textvariable=self.var_arch_to, width=11).pack(side='left', padx=(2, 8))
        ttk.Label(arch_filters, text='Agency:').pack(side='left')
        self.var_arch_agency = _var('')
        self.agency_box = ttk.Combobox(arch_filters, textvariable=self.var_arch_agency,
                                       width=16, state='readonly', values=('',))
        self.agency_box.pack(side='left', padx=(2, 8))
        ttk.Label(arch_filters, text='Action:').pack(side='left')
        self.var_arch_action = _var('')
        self.action_box = ttk.Combobox(arch_filters, textvariable=self.var_arch_action,
                                       width=13, state='readonly', values=('',))
        self.action_box.pack(side='left', padx=(2, 8))
        ttk.Label(arch_filters, text='Attachments:').pack(side='left')
        self.var_arch_attach = _var('Any')
        self.attach_box = ttk.Combobox(arch_filters, textvariable=self.var_arch_attach,
                                       width=6, state='readonly', values=('Any', 'Yes', 'No'))
        self.attach_box.pack(side='left', padx=(2, 8))
        ttk.Button(arch_filters, text='Clear filters', command=self.clear_archive_filters).pack(side='left')
        for var in (self.var_arch_search, self.var_arch_from, self.var_arch_to,
                    self.var_arch_agency, self.var_arch_action, self.var_arch_attach):
            var.trace_add('write', lambda *args: self.populate_archive())
        arch_paned = ttk.PanedWindow(archive_frame, orient='vertical')
        arch_paned.grid(row=2, column=0, columnspan=2, sticky='nsew')
        archive_frame.rowconfigure(2, weight=1)
        archive_frame.columnconfigure(0, weight=1)

        table_pane = ttk.Frame(arch_paned)
        table_pane.rowconfigure(0, weight=1)
        table_pane.columnconfigure(0, weight=1)
        arch_paned.add(table_pane, weight=1)

        self.archive_tree = ttk.Treeview(
            table_pane, height=5,
            columns=('ref', 'date', 'agency', 'sender', 'subject',
                     'action', 'response', 'attachments'),
            show='headings', selectmode='browse')
        for key, text, width, stretch in (
                ('ref', 'Ref No.', 110, False), ('date', 'Date', 90, False),
                ('agency', 'Agency', 150, False), ('sender', 'Sender', 180, False),
                ('subject', 'Subject', 320, True), ('action', 'Action', 120, False),
                ('response', 'Response', 100, False), ('attachments', 'Attachments', 160, False)):
            self.archive_tree.heading(
                key, text=text, command=lambda k=key: self.sort_archive(k))
            self.archive_tree.column(key, width=width, anchor='w', stretch=stretch)
        arch_y = ttk.Scrollbar(table_pane, orient='vertical', command=self.archive_tree.yview)
        self.archive_tree.configure(yscrollcommand=arch_y.set)
        self.archive_tree.grid(row=0, column=0, sticky='nsew')
        arch_y.grid(row=0, column=1, sticky='ns')
        self.archive_tree.bind('<Double-1>', lambda event: self.open_archive_pdf())
        self.archive_tree.bind('<<TreeviewSelect>>', lambda event: self.show_archive_detail())

        detail_pane = ttk.Frame(arch_paned)
        detail_pane.rowconfigure(2, weight=1)
        detail_pane.columnconfigure(0, weight=1)
        arch_paned.add(detail_pane, weight=1)

        self.var_edit_ref = _var('')
        self.var_edit_subject = _var('')
        self.var_edit_sender = _var('')
        self.var_edit_recipients = _var('')
        self.var_edit_cc = _var('')
        self.var_edit_trans = _var('')
        self.var_edit_attach = _var('')
        self.var_edit_agency = _var('')
        self.var_edit_action = _var('')
        self.var_edit_response = _var('')
        self.var_edit_notes = _var('')

        fields = ttk.Frame(detail_pane)
        fields.grid(row=0, column=0, sticky='we', pady=(4, 0))
        fields.columnconfigure(0, weight=1)
        fields.columnconfigure(1, weight=1)
        form = ttk.LabelFrame(fields, text='Envelope')
        form.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        form.columnconfigure(1, weight=1)
        tracking = ttk.LabelFrame(fields, text='Tracking')
        tracking.grid(row=0, column=1, sticky='nsew')
        tracking.columnconfigure(1, weight=1)

        row = 0
        for label, var in (('Ref No.', self.var_edit_ref), ('Subject', self.var_edit_subject),
                           ('Transaction Date', self.var_edit_trans), ('Sender', self.var_edit_sender),
                           ('Recipients', self.var_edit_recipients), ('CC', self.var_edit_cc),
                           ('Attachments', self.var_edit_attach)):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky='nw', pady=1)
            ttk.Label(form, textvariable=var, wraplength=230, justify='left').grid(
                row=row, column=1, sticky='w', pady=1)
            row += 1

        ttk.Label(tracking, text='Agency').grid(row=0, column=0, sticky='w', pady=1)
        self.arch_agency_edit = ttk.Combobox(tracking, textvariable=self.var_edit_agency, width=26, values=('',))
        self.arch_agency_edit.grid(row=0, column=1, sticky='we', pady=1)
        ttk.Label(tracking, text='Action').grid(row=1, column=0, sticky='w', pady=1)
        self.arch_action_edit = ttk.Combobox(tracking, textvariable=self.var_edit_action, width=26, values=('',))
        self.arch_action_edit.grid(row=1, column=1, sticky='we', pady=1)
        ttk.Label(tracking, text='Response Date').grid(row=2, column=0, sticky='w', pady=1)
        ttk.Entry(tracking, textvariable=self.var_edit_response,
                  width=28).grid(row=2, column=1, sticky='we', pady=1)
        ttk.Label(tracking, text='Notes').grid(row=3, column=0, sticky='w', pady=1)
        ttk.Entry(tracking, textvariable=self.var_edit_notes,
                  width=28).grid(row=3, column=1, sticky='we', pady=1)

        ttk.Label(detail_pane, text="Message (the sender's own text)").grid(
            row=1, column=0, sticky='w', pady=(6, 0))
        self.arch_detail = tk_text(detail_pane)
        self.arch_detail.grid(row=2, column=0, sticky='nsew')
        self.arch_detail.configure(bg=THEME['field'], fg=THEME['text'], insertbackground=THEME['text'],
                                   relief='flat', padx=8, pady=6, font=(self.ui_font, 9), height=7)
        self.arch_detail.configure(state='disabled')

        arch_actions = ttk.Frame(archive_frame)
        arch_actions.grid(row=3, column=0, columnspan=2, sticky='we', pady=(3, 0))
        ttk.Button(arch_actions, text='Save changes', command=self.save_archive_detail).pack(side='left', padx=3)
        ttk.Button(arch_actions, text='Revert', command=self.show_archive_detail).pack(side='left', padx=3)
        ttk.Button(arch_actions, text='Open PDF', command=self.open_archive_pdf).pack(side='left', padx=3)
        ttk.Button(arch_actions, text='Open EML', command=self.open_archive_eml).pack(side='left', padx=3)
        ttk.Button(arch_actions, text='Open attachments folder',
                   command=lambda: _open_path(os.path.join(self.folder, 'Attachments'))).pack(side='left', padx=3)
        ttk.Button(arch_actions, text='Export view',
                   command=self.export_archive_view).pack(side='left', padx=3)
        ttk.Button(arch_actions, text='Regenerate PDFs',
                   command=self.regenerate_pdfs).pack(side='left', padx=3)

        dash_frame = ttk.Frame(notes)
        notes.add(dash_frame, text='Dashboard')
        dash_bar = ttk.Frame(dash_frame)
        dash_bar.grid(row=0, column=0, columnspan=2, sticky='we', pady=(0, 3))
        ttk.Button(dash_bar, text='Refresh', command=self.refresh_dashboard).pack(side='left')
        self.dash_summary = _var('')
        ttk.Label(dash_bar, textvariable=self.dash_summary).pack(side='left', padx=8)
        ttk.Label(dash_frame, text='By Agency').grid(row=1, column=0, sticky='w', pady=(4, 0))
        ttk.Label(dash_frame, text='By Month').grid(row=1, column=1, sticky='w', pady=(4, 0))
        self.dash_agency = ttk.Treeview(dash_frame, columns=('agency', 'count'),
                                        show='headings', height=6)
        self.dash_agency.heading('agency', text='Agency')
        self.dash_agency.heading('count', text='Count')
        self.dash_agency.column('agency', width=220, stretch=True)
        self.dash_agency.column('count', width=70, anchor='e', stretch=False)
        self.dash_agency.grid(row=2, column=0, sticky='nsew', padx=(0, 6))
        self.dash_month = ttk.Treeview(dash_frame, columns=('month', 'count'),
                                       show='headings', height=6)
        self.dash_month.heading('month', text='Month')
        self.dash_month.heading('count', text='Count')
        self.dash_month.column('month', width=160, stretch=True)
        self.dash_month.column('count', width=70, anchor='e', stretch=False)
        self.dash_month.grid(row=2, column=1, sticky='nsew')
        ttk.Label(dash_frame, text='Outstanding actions (no Action / Response Date yet)').grid(
            row=3, column=0, columnspan=2, sticky='w', pady=(6, 0))
        out_frame = ttk.Frame(dash_frame)
        out_frame.grid(row=4, column=0, columnspan=2, sticky='nsew', pady=(2, 0))
        out_frame.rowconfigure(0, weight=1)
        out_frame.columnconfigure(0, weight=1)
        self.dash_outstanding = ttk.Treeview(
            out_frame, columns=('ref', 'agency', 'subject', 'action', 'response'), show='headings')
        for key, text, width, stretch in (
                ('ref', 'Ref No.', 110, False), ('agency', 'Agency', 160, False),
                ('subject', 'Subject', 380, True), ('action', 'Action', 110, False),
                ('response', 'Response', 100, False)):
            self.dash_outstanding.heading(key, text=text)
            self.dash_outstanding.column(key, width=width, anchor='w', stretch=stretch)
        dash_y = ttk.Scrollbar(out_frame, orient='vertical', command=self.dash_outstanding.yview)
        self.dash_outstanding.configure(yscrollcommand=dash_y.set)
        self.dash_outstanding.grid(row=0, column=0, sticky='nsew')
        dash_y.grid(row=0, column=1, sticky='ns')
        dash_frame.rowconfigure(4, weight=1)
        dash_frame.columnconfigure(0, weight=1)
        dash_frame.columnconfigure(1, weight=1)

        log_frame = ttk.Frame(notes)
        notes.add(log_frame, text='Log')
        self.log_text = tk_text(log_frame)
        self.log_text.grid(row=0, column=0, sticky='nsew')
        log_y = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_y.set)
        log_y.grid(row=0, column=1, sticky='ns')
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text.configure(bg=THEME['field'], fg=THEME['text'], insertbackground=THEME['text'],
                                relief='flat', padx=8, pady=6, font=(self.ui_font, 9))

        self.refresh_keyword_list()
        self.refresh_tree()
        self.refresh_archive()

    # ----------------------------------------------------------- log/status
    def log(self, text):
        self.queue.put(('log', text))

    def status(self, text):
        self.status_var.set(str(text))

    def clear_log(self):
        self.log_text.delete('1.0', 'end')

    # ------------------------------------------------------------ settings
    def open_settings(self):
        self.dialog = SettingsDialog(self)

    def toggle_theme(self):
        dark = not self.dark
        if not self._set_theme(dark):
            return
        self.log('Dark mode is now ' + ('on.' if dark else 'off.'))

    def _set_theme(self, dark):
        self.dark = bool(dark)
        self.config['dark_mode'] = self.dark
        try:
            archive.save_config(self.config_path, self.config)
        except OSError as exc:
            self.status(f'Could not save the theme: {exc}')
            return False
        self.ui_font = apply_theme(self.root_window, self.dark)
        self._restyle_widgets()
        if hasattr(self, 'btn_theme'):
            self.btn_theme.configure(
                text='Use light mode' if self.dark else 'Use dark mode')
        dialog = getattr(self, 'dialog', None)
        if dialog is not None and dialog.window.winfo_exists():
            dialog.window.destroy()
            self.open_settings()
        self.refresh_tree()
        return True

    def _restyle_widgets(self):
        if hasattr(self, 'log_text'):
            self.log_text.configure(bg=THEME['field'], fg=THEME['text'],
                                    insertbackground=THEME['text'])
        if hasattr(self, 'keywords'):
            self.keywords.configure(bg=THEME['field'], fg=THEME['text'],
                                    selectbackground=THEME['accent'],
                                    selectforeground='#ffffff')
        if hasattr(self, 'tree'):
            self.tree.tag_configure(
                'thread', background=THEME['thread_bg'], foreground=THEME['heading'],
                font=(self.ui_font, 9, 'bold'))
            self.tree.tag_configure('solo', font=(self.ui_font, 9, 'bold'))
            self.tree.tag_configure('unread', font=(self.ui_font, 9, 'bold'),
                                    foreground=THEME['danger'])
            self.tree.tag_configure('alt', background=THEME['alt'])

    def open_manual(self):
        manual = os.path.join(archive.SCRIPT_DIR, MANUAL_FILE)
        if os.path.isfile(manual):
            _open_path(manual)
        else:
            _messagebox(
                'info', 'Manual not found',
                f'The manual "{MANUAL_FILE}" should be saved next to the program '
                f'in ( {archive.SCRIPT_DIR} ).\n\nCopy the PDF there (or next to the .exe) '
                'and press this button again.')

    def uninstall(self):
        import tkinter.messagebox as _mb
        root = archive.SCRIPT_DIR
        answer = _mb.askyesnocancel(
            'Uninstall ' + APP_NAME,
            'Remove all ' + APP_NAME + ' data from this computer?\n\n'
            'This deletes your settings, sign-in cache, log and the Desktop shortcut, '
            'in/from:\n' + root + '\n\n'
            'YES = also delete your archived emails (PDF, EML, attachments).\n'
            'NO  = keep your archived emails, remove everything else.\n'
            'CANCEL = do nothing.')
        if answer is None:
            return
        removed = archive.remove_local_data(bool(answer))
        manual = os.path.join(archive.SCRIPT_DIR, MANUAL_FILE)
        if os.path.isfile(manual):
            try:
                os.remove(manual)
                removed.append(manual)
            except OSError:
                pass
        shortcut = os.path.join(_desktop_dir(), APP_NAME + '.lnk')
        if os.path.isfile(shortcut):
            try:
                os.remove(shortcut)
                removed.append(shortcut)
            except OSError:
                pass
        scheduled = archive.schedule_self_removal()
        kept = ('' if answer else '\n\nYour archived emails were kept.')
        tail = ('\n\nThe program file will be removed automatically when the window closes.'
                if scheduled else '\n\nDelete the program file (' + APP_NAME + '.exe) yourself.')
        _messagebox(
            'info', 'Uninstalled',
            f'Removed {len(removed)} item(s).' + kept + tail)
        self.destroy()

    def about(self):
        _messagebox(
            'info', 'About ' + APP_NAME,
            APP_NAME + '\n\nArchives Outlook, Gmail, Yahoo or any IMAP inbox into PDF, '
            'EML, attachments and a CSV/Excel log.\n\nUses Microsoft OAuth2 (Outlook) or '
            'app-password sign-in (Gmail/Yahoo).\n\nManual: use the "Open PDF Manual" '
            'button on the left sidebar.\n\n'
            'Copyright \u00a9 2026 Eng. Nasser Taher. All rights reserved.')

    def apply_settings(self, updates):
        errors = []
        email = str(updates.get('email') or '').strip()
        if not email:
            errors.append('Email address cannot be empty.')
        provider = str(updates.get('imap_provider') or 'custom')
        if provider not in GUI_PROVIDERS:
            errors.append('Unknown mail provider - choose one from the list.')
        oauth = provider == 'outlook'
        client_id = str(updates.get('client_id') or '').strip()
        if oauth and not client_id:
            errors.append('Outlook.com needs the Azure app (client) ID - see the link in the '
                          'Settings window.')
        tenant = str(updates.get('tenant') or '').strip() or 'common'
        if tenant not in ('common', 'consumers', 'organizations'):
            errors.append(f'Tenant must be common, consumers or organizations - got "{tenant}".')
        server = str(updates.get('imap_server') or '').strip()
        if not server:
            errors.append('IMAP server cannot be empty.')
        try:
            port = int(updates.get('imap_port'))
            if not 1 <= port <= 65535:
                raise ValueError
        except (TypeError, ValueError):
            errors.append('IMAP port must be a number between 1 and 65535.')
        try:
            limit = int(updates.get('fetch_limit'))
            if not 1 <= limit <= 5000:
                raise ValueError
        except (TypeError, ValueError):
            errors.append('Fetch limit must be a number between 1 and 5000.')
        if errors:
            _messagebox('warning', 'Check your settings', '\n'.join(errors))
            return False
        app_password = str(updates.get('app_password') or '').strip()
        if provider in ('gmail', 'yahoo') and not app_password:
            _messagebox('warning', 'App password missing',
                        'Gmail/Yahoo will refuse the normal password - create an app password '
                        '(link in the Settings window) and paste it in the App password field.')
        sign_in_changed = (
            self.config.get('email') != email
            or self.config.get('client_id') != client_id
        )
        self.config['email'] = email
        self.config['imap_provider'] = provider
        if not oauth:
            self.config['app_password'] = app_password
        self.config['client_id'] = client_id
        self.config['tenant'] = tenant
        self.config['imap_server'] = server
        self.config['imap_port'] = port
        self.config['use_ssl'] = bool(updates.get('use_ssl'))
        self.config['fetch_limit'] = limit
        self.config['thread_view'] = bool(updates.get('thread_view'))
        try:
            archive.save_config(self.config_path, self.config)
        except OSError as exc:
            _messagebox('error', 'Could not save settings', str(exc))
            return False
        self.var_email.set(email)
        self.status('Settings saved.')
        self.log('Settings saved.')
        if sign_in_changed:
            self.log('The Microsoft sign-in is for the old account/app - the next fetch will ask '
                     'you to sign in again, or delete token_cache.json to force it now.')
        if self.worker is None or not self.worker.is_alive():
            self.refresh_tree()
        return True

    # --------------------------------------------------------------- fetch
    def fetch(self):
        if not self.config.get('client_id') or not self.config.get('email'):
            self.log('Tip: set "client_id" in config.json (OAuth2) - see README section 1.')
        self.run_worker(self._worker_fetch)

    def _worker_fetch(self):
        self.queue.put(('log', f'Connecting to {self.config.get("imap_server")}...'))
        result = archive.cmd_fetch(self.config, self.folder)
        self.queue.put(('fetched', result or []))

    # ------------------------------------------------------------- archive
    def archive(self):
        positions = selected_positions(self.summaries, self.selected_keys)
        if not positions:
            self.status('Nothing selected to archive - click the Keep column.')
            return
        snapshot = list(self.summaries)
        self.run_worker(lambda: self._worker_archive(snapshot, positions))

    def _worker_archive(self, summaries, positions):
        config = dict(archive.load_config(self.config_path))
        root = archive.resolve_root(config)
        archive.ensure_dirs(root)

        def notify(text):
            self.queue.put(('log', text))

        archive.archive_selected(config, root, summaries, positions, notify)
        self.queue.put(('archive-refresh', archive.load_seen(root)))

    # ------------------------------------------------------------- export
    def export_excel(self):
        self.run_worker(self._worker_export)

    def _worker_export(self):
        archive.export_excel(self.folder)
        self.queue.put(('log', 'Excel log exported: log.xlsx'))
        self.queue.put(('status', 'Excel log exported: log.xlsx'))

    # ------------------------------------------------------------ keywords
    def remove_keyword(self):
        selection = self.keywords.curselection()
        if not selection:
            return
        chosen = self.keywords.get(selection[0])
        self.hidden_names.discard(chosen)
        self.apply_keywords()

    def clear_keywords(self):
        if self.hidden_names:
            self.hidden_names.clear()
            self.apply_keywords()

    def apply_keywords(self):
        words = reconcile_keywords(self.config.get('exclude_keywords') or [],
                                   self.summaries, self.hidden_names)
        self.config['exclude_keywords'] = words
        try:
            archive.save_config(self.config_path, self.config)
        except OSError as exc:
            self.status(f'Could not save keywords: {exc}')
            return
        self.refresh_keyword_list()
        self.status(f'Keywords saved ({len(words)}).')
        self.log(f'Keywords saved to {os.path.basename(self.config_path)} ({len(words)}).')
        self.refresh_tree()

    def refresh_keyword_list(self):
        self.keywords.delete(0, 'end')
        for word in self.config.get('exclude_keywords') or []:
            self.keywords.insert('end', word)

    # ---------------------------------------------------------- archive tab
    def refresh_archive(self):
        self.log_rows = archive.read_log_rows(self.folder)
        agencies = sorted({str(row.get('Agency') or '').strip() for row in self.log_rows
                           if str(row.get('Agency') or '').strip() not in ('', archive.DASH)})
        actions = sorted({str(row.get('Action') or '').strip() for row in self.log_rows
                          if str(row.get('Action') or '').strip() not in ('', archive.DASH)})
        self.agency_box['values'] = ('',) + tuple(agencies)
        self.action_box['values'] = ('',) + tuple(actions)
        self.arch_agency_edit['values'] = ('',) + tuple(agencies)
        self.arch_action_edit['values'] = ('',) + tuple(actions)
        if self.var_arch_agency.get() not in self.agency_box['values']:
            self.var_arch_agency.set('')
        if self.var_arch_action.get() not in self.action_box['values']:
            self.var_arch_action.set('')
        self.populate_archive()
        self.refresh_dashboard()

    def refresh_dashboard(self):
        stats = archive.archive_stats(getattr(self, 'log_rows', []))
        self.dash_summary.set(
            f"{stats['total']} record(s) - {stats['threads']} thread(s) - "
            f"{stats['with_attachments']} with attachments - "
            f"{len(stats['outstanding'])} outstanding")
        self.dash_agency.delete(*self.dash_agency.get_children())
        for agency, count in sorted(stats['agencies'].items(), key=lambda kv: (-kv[1], kv[0])):
            self.dash_agency.insert('', 'end', values=(agency, count))
        self.dash_month.delete(*self.dash_month.get_children())
        for month, count in sorted(stats['months'].items()):
            self.dash_month.insert('', 'end', values=(month, count))
        self.dash_outstanding.delete(*self.dash_outstanding.get_children())
        for row in stats['outstanding']:
            self.dash_outstanding.insert('', 'end', values=(
                row.get('Ref No.', ''), row.get('Agency', ''), row.get('Subject', ''),
                row.get('Action', ''), row.get('Response Date', '')))

    def export_archive_view(self):
        from tkinter import filedialog
        rows = getattr(self, 'archive_filtered', [])
        if not rows:
            self.status('Nothing to export - no rows match the current filters.')
            return
        path = filedialog.asksaveasfilename(
            title='Export the filtered archive view',
            defaultextension='.xlsx',
            filetypes=[('Excel workbook', '*.xlsx'), ('CSV file', '*.csv')])
        if not path:
            return
        try:
            if path.lower().endswith('.csv'):
                archive.export_rows_csv(path, rows)
            else:
                archive.export_rows_excel(path, rows)
        except (OSError, archive.ArchiveError) as exc:
            self.status(f'Export failed: {exc}')
            return
        self.status(f'Exported {len(rows)} row(s) to {path}')

    def regenerate_pdfs(self):
        if not self.log_rows:
            self.status('Archive log is empty - nothing to regenerate.')
            return
        if not _messagebox('ask', 'Regenerate PDFs',
                           'Rebuild every archived PDF from the saved .eml files?\n'
                           'This overwrites the PDFs in the PDF folder and removes the '
                           'OneDrive/Outlook link clutter from the message bodies.'):
            return
        self.run_worker(self._worker_rebuild_pdfs)

    def _worker_rebuild_pdfs(self):
        result = archive.rebuild_pdfs(self.folder, lambda text: self.queue.put(('log', text)))
        self.queue.put(('log', f'Rebuilt {result["rebuilt"]} PDF(s), skipped {result["skipped"]}.'))
        self.queue.put(('status', f'Rebuilt {result["rebuilt"]} PDF(s) from saved emails.'))

    def clear_archive_filters(self):
        self.var_arch_search.set('')
        self.var_arch_from.set('')
        self.var_arch_to.set('')
        self.var_arch_agency.set('')
        self.var_arch_action.set('')
        self.var_arch_attach.set('Any')

    def import_archive_csv(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title='Import CSV into the archive log',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
        if not path:
            return
        try:
            incoming = archive.parse_import_csv(path)
        except (OSError, archive.ArchiveError) as exc:
            self.status(f'Import failed: {exc}')
            return
        if not incoming:
            self.status('No rows found in that CSV.')
            return
        merged, added, skipped = archive.merge_log_rows(self.log_rows, incoming)
        try:
            archive.write_log_rows(self.folder, merged)
        except OSError as exc:
            self.status(f'Import failed while saving: {exc}')
            return
        self.refresh_archive()
        self.status(f'Imported {added} row(s), skipped {skipped} duplicate(s).')

    def populate_archive(self):
        self.archive_tree.delete(*self.archive_tree.get_children())
        rows = getattr(self, 'log_rows', [])
        attach = {'Yes': True, 'No': False}.get(self.var_arch_attach.get())
        shown = archive.filter_log_rows(
            rows, text=self.var_arch_search.get(), date_from=self.var_arch_from.get(),
            date_to=self.var_arch_to.get(), agency=self.var_arch_agency.get(),
            action=self.var_arch_action.get(), attachment=attach)
        if self.arch_sort_key:
            if self.arch_sort_key == 'sender':
                shown.sort(key=lambda row: (row.get('Sender Name') or row.get('Sender Email') or '').casefold(),
                           reverse=self.arch_sort_desc)
            else:
                column = ARCHIVE_COLUMNS.get(self.arch_sort_key)
                shown.sort(key=lambda row: archive.log_row_sort_key(row, column),
                           reverse=self.arch_sort_desc)
        self.archive_filtered = shown
        for index, row in enumerate(shown):
            self.archive_tree.insert('', 'end', iid='log-%d' % (index + 1), values=(
                row.get('Ref No.', ''), row.get('Transaction Date', ''),
                row.get('Agency', ''), row.get('Sender Name', '') or row.get('Sender Email', ''),
                row.get('Subject', ''), row.get('Action', ''),
                row.get('Response Date', ''), row.get('Attachments', '')))
        if len(shown) == len(rows):
            self.arch_count.set(f'{len(rows)} archived record(s).')
        else:
            self.arch_count.set(f'showing {len(shown)} of {len(rows)}')

    def _selected_archive_row(self):
        selection = self.archive_tree.selection()
        if not selection:
            return None
        prefix = ''.join(ch for ch in selection[0] if ch.isdigit())
        if not prefix:
            return None
        rows = getattr(self, 'archive_filtered', [])
        index = int(prefix) - 1
        if 0 <= index < len(rows):
            return rows[index]
        return None

    def _archive_eml_path(self, ref, row=None):
        if not ref:
            return ''
        eml_dir = os.path.join(self.folder, 'EML')
        if not os.path.isdir(eml_dir):
            return ''
        candidates = []
        for name in os.listdir(eml_dir):
            if name == ref + '.eml':
                candidates.append((1, 0, name))
            elif name.startswith(ref + '-'):
                numbers = [int(part) for part in re.findall(r'\d+', name[len(ref):])]
                first = numbers[0] if numbers else 1
                second = numbers[1] if len(numbers) > 1 else 0
                candidates.append((first, second, name))
        candidates.sort()
        paths = [os.path.join(eml_dir, name) for _, _, name in candidates]
        if not paths:
            return ''
        if row is not None:
            position = 0
            for candidate in self.log_rows:
                if str(candidate.get('Ref No.') or '').strip() != ref:
                    continue
                if candidate is row:
                    return paths[position] if position < len(paths) else paths[0]
                position += 1
        return paths[0]

    def show_archive_detail(self):
        row = self._selected_archive_row()
        if row is None:
            self._clear_archive_form()
            self._set_archive_detail('')
            return
        self.var_edit_ref.set(str(row.get('Ref No.') or ''))
        self.var_edit_subject.set(str(row.get('Subject') or ''))
        self.var_edit_sender.set(str(row.get('Sender Name') or '') or str(row.get('Sender Email') or ''))
        recipients = str(row.get('Recipients') or '')
        self.var_edit_recipients.set('' if recipients == archive.DASH else recipients)
        cc = str(row.get('CC') or '')
        self.var_edit_cc.set('' if cc == archive.DASH else cc)
        self.var_edit_trans.set(str(row.get('Transaction Date') or ''))
        self.var_edit_attach.set(str(row.get('Attachments') or ''))
        agency = str(row.get('Agency') or '')
        self.var_edit_agency.set('' if agency == archive.DASH else agency)
        placeholders = (archive.DASH, archive.FILL_MANUALLY)
        action = str(row.get('Action') or '')
        self.var_edit_action.set('' if action in placeholders else action)
        response = str(row.get('Response Date') or '')
        self.var_edit_response.set('' if response in placeholders else response)
        notes = str(row.get('Notes') or '')
        self.var_edit_notes.set('' if notes == archive.DASH else notes)
        ref = str(row.get('Ref No.') or '').strip()
        eml = self._archive_eml_path(ref, row)
        body = archive.strip_quoted_reply(archive.eml_body(eml)) if eml else '(no saved .eml for this record)'
        self._set_archive_detail(body)

    def _clear_archive_form(self):
        for var in (self.var_edit_ref, self.var_edit_subject, self.var_edit_sender,
                    self.var_edit_recipients, self.var_edit_cc,
                    self.var_edit_trans, self.var_edit_attach, self.var_edit_agency,
                    self.var_edit_action, self.var_edit_response, self.var_edit_notes):
            var.set('')

    def save_archive_detail(self):
        row = self._selected_archive_row()
        if row is None:
            self.status('Select an archived record first.')
            return
        row['Agency'] = self.var_edit_agency.get().strip() or archive.DASH
        row['Action'] = self.var_edit_action.get().strip() or archive.DASH
        row['Response Date'] = self.var_edit_response.get().strip() or archive.DASH
        row['Notes'] = self.var_edit_notes.get().strip() or archive.DASH
        ref = str(row.get('Ref No.') or '').strip()
        try:
            archive.write_log_rows(self.folder, self.log_rows)
        except OSError as exc:
            self.status(f'Save failed: {exc}')
            return
        self.refresh_archive()
        self._reselect_archive(ref)
        self.status(f'{ref}: saved to log.csv and log.xlsx')

    def _reselect_archive(self, ref):
        for index, row in enumerate(getattr(self, 'archive_filtered', [])):
            if str(row.get('Ref No.') or '').strip() == ref:
                iid = 'log-%d' % (index + 1)
                if self.archive_tree.exists(iid):
                    self.archive_tree.selection_set(iid)
                    self.archive_tree.see(iid)
                return

    def _set_archive_detail(self, text):
        self.arch_detail.configure(state='normal')
        self.arch_detail.delete('1.0', 'end')
        self.arch_detail.insert('1.0', text)
        self.arch_detail.configure(state='disabled')

    def open_archive_pdf(self):
        row = self._selected_archive_row()
        if row is None:
            return
        ref = str(row.get('Ref No.') or '').strip()
        if not ref:
            self.status('This record has no Ref No.')
            return
        pdf = os.path.join(self.folder, 'PDF', ref + '.pdf')
        if os.path.exists(pdf):
            _open_path(pdf)
        else:
            self.status(f'PDF not found: {pdf}')

    def open_archive_eml(self):
        row = self._selected_archive_row()
        if row is None:
            self.status('Select an archived record first.')
            return
        ref = str(row.get('Ref No.') or '').strip()
        eml = self._archive_eml_path(ref, row)
        if eml:
            _open_path(eml)
        else:
            self.status(f'No .eml saved for {ref or "this record"}.')

    def sort_archive(self, key):
        column = ARCHIVE_COLUMNS.get(key)
        rows = getattr(self, 'log_rows', [])
        if not column or not rows:
            return
        if self.arch_sort_key == key:
            self.arch_sort_desc = not self.arch_sort_desc
        else:
            self.arch_sort_key = key
            self.arch_sort_desc = False
        self.populate_archive()
        self._update_archive_headings()

    def _update_archive_headings(self):
        for key, text in ARCHIVE_HEADINGS.items():
            arrow = ''
            if key == self.arch_sort_key:
                arrow = ' \u25bc' if self.arch_sort_desc else ' \u25b2'
            self.archive_tree.heading(key, text=text + arrow)

    # ----------------------------------------------------------- tree
    def clear_email_filters(self):
        self.var_filter.set('')
        self.var_email_from.set('')
        self.var_email_to.set('')
        self.var_unread_only.set(False)
        self.var_has_attach.set(False)
        self.var_archived.set('Any')

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.rowmap = {}
        archived = {'Archived': True, 'Not archived': False}.get(self.var_archived.get())
        rows = ui_rows(self.summaries, self.selected_keys, self.hidden_names,
                       self.config.get('thread_view', True), self.var_filter.get(),
                       date_from=self.var_email_from.get(), date_to=self.var_email_to.get(),
                       unread_only=bool(self.var_unread_only.get()),
                       has_attachment=bool(self.var_has_attach.get()),
                       archived=archived)
        mail_on_even = 0
        for row in rows:
            if row['kind'] == 'thread':
                iid = row['iid']
                self.rowmap[iid] = ('thread', row['thread_no'])
                self.tree.insert('', 'end', iid=iid, values=(
                    '', '', 'Thread %d' % row['thread_no'], '%d messages' % row['count'], ''),
                    tags=('thread',))
            else:
                iid = row['iid']
                self.rowmap[iid] = ('mail', row['index'])
                mark_keep = _mark(row['selected'], row['archived'])
                mark_hide = _mark(row['hidden'], False)
                tags = []
                if mail_on_even % 2:
                    tags.append('alt')
                mail_on_even += 1
                if not row['thread_no']:
                    tags.append('solo')
                if row['unread']:
                    tags.append('unread')
                self.tree.insert('', 'end', iid=iid, values=(
                    mark_keep, mark_hide, row['date'], row['sender'], row['subject']),
                    tags=tuple(tags))

    def on_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        iid = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        row = self.rowmap.get(iid)
        if not row:
            return
        kind, key = row
        if kind == 'thread':
            if column == '#1':
                self.toggle_thread(key)
            return
        item = self.summaries[key]
        if column == '#1':
            self.toggle_thread(item.get('thread_no') or 0, fallback_key=item['key'])
        elif column == '#2':
            self.toggle_hide(item)

    def toggle_thread(self, thread_no, fallback_key=None):
        if thread_no:
            members = [item for item in self.summaries
                       if (item.get('thread_no') or 0) == thread_no]
        elif fallback_key:
            members = [item for item in self.summaries if item['key'] == fallback_key]
        else:
            return
        keys = [item['key'] for item in members]
        turn_on = not all(key in self.selected_keys for key in keys)
        if turn_on:
            self.selected_keys.update(keys)
        else:
            for key in keys:
                self.selected_keys.discard(key)
        self.refresh_tree()
        if turn_on:
            self.status(f'Marked {len(keys)} message(s) to archive - press Archive selected.')
        else:
            self.status('Selection cleared.')

    def toggle_hide(self, item):
        keywords_list = archive.keywords_for(item)
        if not keywords_list:
            self.status('This email has no sender name or address to filter on.')
            return
        any_hidden = any(keyword in self.hidden_names for keyword in keywords_list)
        for keyword in keywords_list:
            if any_hidden:
                self.hidden_names.discard(keyword)
            else:
                self.hidden_names.add(keyword)
        self.apply_keywords()

    # ------------------------------------------------------------- worker
    def run_worker(self, job):
        if self.worker and self.worker.is_alive():
            self.status('Another task is still running - waiting for it to finish.')
            return
        self.status('Working...')
        self.worker = threading.Thread(target=self._thread_wrapper, args=(job,), daemon=True)
        self.worker.start()

    def _thread_wrapper(self, job):
        try:
            job()
        except archive.ArchiveError as exc:
            self.queue.put(('error', str(exc)))
        except Exception as exc:
            detail = str(exc) + '\n' + traceback.format_exc(limit=4)
            self.queue.put(('error', detail))

    # ------------------------------------------------------------- poll
    def poll(self):
        try:
            while True:
                event, payload = self.queue.get_nowait()
                self.handle_event(event, payload)
        except queue.Empty:
            pass
        self.root_window.after(self.TICK, self.poll)

    def handle_event(self, event, payload):
        if event == 'log':
            self.log_text.insert('end', payload + '\n')
            self.log_text.see('end')
        elif event == 'status':
            self.status(payload)
        elif event == 'fetched':
            self.summaries = payload
            self.selected_keys = set()
            self.seen = archive.load_seen(self.folder)
            self.refresh_tree()
            self.status(f'{len(self.summaries)} email(s) shown.')
        elif event == 'error':
            self.log('[error] ' + str(payload))
            self.status('Task failed - see the log.')
            _messagebox('error', 'Task failed', str(payload))
        elif event == 'archive-refresh':
            self.seen = payload
            for item in self.summaries:
                item['archived'] = item['key'] in self.seen
            self.refresh_tree()
            self.refresh_archive()
            self.status('Archive finished - see the log.')

    def on_close(self):
        if self.worker and self.worker.is_alive():
            _messagebox('info', 'Still working',
                        'A fetch/archive task is still running in the background.\n'
                        'It will finish before the window closes.')
            return
        self.root_window.destroy()


def _var(value):
    import tkinter as _tk
    return _tk.StringVar(value=str(value))


def _link(parent, text, url):
    import tkinter as tk
    label = tk.Label(parent, text=text, fg=THEME['link'], bg=THEME['bg'], cursor='hand2',
                     font=('TkDefaultFont', 9, 'underline'))

    def activate(_event=None):
        webbrowser.open(url)

    label.bind('<Button-1>', activate)
    label.url = url
    label.activate = activate
    return label


def _mark(checked, archived):
    if archived:
        return '\u2714'
    return '\u2611' if checked else '\u2610'


def _open_path(path):
    try:
        os.startfile(path)
    except Exception as exc:
        _messagebox('error', 'Cannot open', str(exc))


def _messagebox(kind, title, message):
    import tkinter.messagebox as _mb
    if kind == 'ask':
        return _mb.askyesno(title, message)
    func = {'info': _mb.showinfo, 'warning': _mb.showwarning, 'error': _mb.showerror}.get(kind, _mb.showinfo)
    func(title, message)


def _desktop_dir():
    try:
        import ctypes
        buffer = ctypes.create_unicode_buffer(260)
        ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buffer)
        if buffer.value:
            return buffer.value
    except Exception:
        pass
    return os.path.join(os.path.expanduser('~'), 'Desktop')


def create_desktop_shortcut():
    if not getattr(sys, 'frozen', False):
        return
    target = os.path.abspath(sys.executable)
    if not target.lower().endswith('.exe'):
        return
    shortcut = os.path.join(_desktop_dir(), APP_NAME + '.lnk')
    if os.path.isfile(shortcut):
        return
    folder = os.path.dirname(target)
    script = (
        'Set s = CreateObject("WScript.Shell")\r\n'
        'Set f = CreateObject("Scripting.FileSystemObject")\r\n'
        'd = s.SpecialFolders("Desktop")\r\n'
        'p = d & "\\\\' + APP_NAME + '.lnk"\r\n'
        'If Not f.FileExists(p) Then\r\n'
        '  Set l = s.CreateShortcut(p)\r\n'
        '  l.TargetPath = "' + target + '"\r\n'
        '  l.WorkingDirectory = "' + folder + '"\r\n'
        '  l.IconLocation = "' + target + ',0"\r\n'
        '  l.Description = "' + APP_NAME + '"\r\n'
        '  l.Save\r\n'
        'End If\r\n'
    )
    import subprocess
    import tempfile
    path = os.path.join(tempfile.gettempdir(), 'email_archive_tool_shortcut.vbs')
    try:
        with open(path, 'w', encoding='utf-8') as handle:
            handle.write(script)
        subprocess.run(['cscript', '//nologo', path],
                       creationflags=0x08000000, timeout=30,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def tk_text(parent):
    from tkinter.scrolledtext import ScrolledText
    return ScrolledText(parent, height=9, wrap='word')


def main(argv=None):
    parser = argparse.ArgumentParser(description='GUI for the Email Archive Tool')
    parser.add_argument('--config', metavar='PATH', default=None,
                        help='path to config.json (default: config.json next to archive.py)')
    parser.add_argument('--selfcheck', action='store_true',
                        help='build the interface, then exit 0 (success) or 1 (failure) without the main loop')
    args = parser.parse_args(argv)
    config_path = os.path.abspath(args.config) if args.config else archive.default_config_path()
    import tkinter as tk
    archive.setup_console()
    if args.selfcheck:
        return _selfcheck(config_path)
    root_window = tk.Tk()
    install_log_hook()
    app = MailToolApp(root_window, config_path)
    if getattr(sys, 'frozen', False):
        create_desktop_shortcut()
    app.root_window.mainloop()
    return 0


def _selfcheck(config_path):
    import json
    import tempfile
    import tkinter as tk
    folder = tempfile.mkdtemp(prefix='archive_selfcheck_')
    try:
        cfg = os.path.join(folder, 'config.json')
        with open(cfg, 'w', encoding='utf-8') as handle:
            json.dump({'email': 'selfcheck@example.com', 'client_id': '', 'archive_folder': folder}, handle)
        root_window = tk.Tk()
        install_log_hook()
        app = MailToolApp(root_window, cfg)
        app.open_settings()
        root_window.update()
        dialog = getattr(app, 'dialog', None)
        if dialog is not None:
            dialog.window.destroy()
        root_window.update()
        root_window.destroy()
        return 0
    except Exception:
        return 1
    finally:
        if os.path.isdir(folder):
            try:
                shutil.rmtree(folder)
            except OSError:
                pass


if __name__ == '__main__':
    sys.exit(main())