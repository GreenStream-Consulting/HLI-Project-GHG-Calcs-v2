import math
import re
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import requests
from openpyxl import load_workbook

APP_TITLE = "HLI Project GHG Calcs Builder"
USER_AGENT = "HLI-Project-GHG-Calcs/1.0 (local desktop app)"
REQUEST_TIMEOUT = 20

INPUT_HEADERS = [
    'Project ID', 'Pick Up Date', 'Weight (Pounds)', 'Origin', 'Destination', 'Mode',
    'Delivery Date', 'Distance (Miles)', 'Subcontractor/Partner', 'Client',
    '(Road) Vehicle category (US)', '(Sea) Vessel Type', '(Inland Water) Vessel Type'
]
COLS_A_TO_N = list(range(1, 15))

ALIASES = {
    'loredo, tx': 'Laredo, TX',
    'loredo tx': 'Laredo, TX',
    'new york city, ny': 'New York, NY',
    'port of freeport': 'Port Freeport, Freeport, TX',
    'port of houston': 'Port of Houston, Houston, TX',
}

PORT_HINTS = {
    'the hague, netherlands': 'Port of Rotterdam, Netherlands',
    'new york, ny': 'Port of New York, NY',
    'new york city, ny': 'Port of New York, NY',
    'los angeles, california': 'Port of Los Angeles, CA',
    'cartagena, colombia': 'Port of Cartagena, Colombia',
    'port of freeport': 'Port Freeport, Freeport, TX',
    'port of houston': 'Port of Houston, Houston, TX',
    'myrtle beach, sc': 'Port of Charleston, SC',
}

_geocode_cache = {}
_route_cache = {}


def normalize_text(value):
    if value is None:
        return ''
    text = str(value).strip()
    lowered = re.sub(r'\s+', ' ', text.lower())
    return ALIASES.get(lowered, text)


def looks_blank_or_zero(value):
    if value is None:
        return True
    if isinstance(value, str):
        text = value.strip()
        return text in {'', '0', '0.0', 'NAME?', '#NAME?'}
    return value == 0


def clean_output_value(value):
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        if text in {'', '0', '0.0', 'NAME?', '#NAME?'}:
            return None
        return text
    if value == 0:
        return None
    return value


def parse_headers(ws, header_row=1):
    mapping = {}
    for c in range(1, ws.max_column + 1):
        header = ws.cell(header_row, c).value
        if header is not None:
            mapping[str(header).strip()] = c
    return mapping


def geocode_location(query):
    query = normalize_text(query)
    if not query:
        return None
    if query in _geocode_cache:
        return _geocode_cache[query]
    try:
        response = requests.get(
            'https://nominatim.openstreetmap.org/search',
            params={'format': 'jsonv2', 'limit': 1, 'q': query},
            headers={'User-Agent': USER_AGENT},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        data = response.json()
        if not data:
            _geocode_cache[query] = None
            return None
        result = {
            'lat': float(data[0]['lat']),
            'lon': float(data[0]['lon']),
            'display_name': data[0]['display_name'],
        }
        _geocode_cache[query] = result
        return result
    except Exception:
        _geocode_cache[query] = None
        return None


def road_distance_miles(origin, destination):
    origin = normalize_text(origin)
    destination = normalize_text(destination)
    key = (origin, destination)
    if key in _route_cache:
        return _route_cache[key]
    a = geocode_location(origin)
    b = geocode_location(destination)
    if not a or not b:
        _route_cache[key] = None
        return None
    try:
        response = requests.get(
            f'https://router.project-osrm.org/route/v1/driving/{a["lon"]},{a["lat"]};{b["lon"]},{b["lat"]}',
            params={'overview': 'false'},
            headers={'User-Agent': USER_AGENT},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        data = response.json()
        routes = data.get('routes') or []
        if not routes:
            _route_cache[key] = None
            return None
        miles = routes[0]['distance'] / 1609.344
        _route_cache[key] = miles
        return miles
    except Exception:
        _route_cache[key] = None
        return None


def haversine_miles(a, b):
    lat1, lon1 = math.radians(a['lat']), math.radians(a['lon'])
    lat2, lon2 = math.radians(b['lat']), math.radians(b['lon'])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 3958.7613 * 2 * math.asin(math.sqrt(h))


def sea_or_water_distance_miles(origin, destination, mode):
    mode_key = (mode or '').strip().lower()
    origin_q = PORT_HINTS.get(normalize_text(origin).lower(), normalize_text(origin))
    destination_q = PORT_HINTS.get(normalize_text(destination).lower(), normalize_text(destination))
    a = geocode_location(origin_q)
    b = geocode_location(destination_q)
    if not a or not b:
        return None, 'Distance estimate unavailable because geocoding failed'
    base = haversine_miles(a, b)
    factor = 1.30 if mode_key == 'sea' else 1.25
    miles = max(1, round(base * factor))
    label = 'sea route' if mode_key == 'sea' else 'inland water route'
    method = f'Estimated using online geocoding and {label} mileage factor'
    if origin_q != normalize_text(origin) or destination_q != normalize_text(destination):
        method += f' ({origin_q} to {destination_q})'
    return miles, method


def estimate_distance(origin, destination, mode):
    mode_key = (mode or '').strip().lower()
    if mode_key == 'storage':
        return None, None
    if mode_key == 'road':
        miles = road_distance_miles(origin, destination)
        if miles is None:
            return None, 'Road estimate unavailable because online routing failed'
        return max(1, round(miles)), 'Estimated using online road routing'
    if mode_key == 'rail':
        miles = road_distance_miles(origin, destination)
        if miles is None:
            return None, 'Rail estimate unavailable because road routing failed'
        return max(1, round(miles * 0.90)), 'Estimated rail distance as 90% of road distance using online routing'
    if mode_key in {'sea', 'inland water'}:
        return sea_or_water_distance_miles(origin, destination, mode)
    return None, 'Distance estimate unavailable because mode was not recognized'


def compatible_formula(row_num, mode):
    r = row_num
    if mode == 'road':
        return f'=IF(F{r}="Road",((C{r}/2204.62)*M{r}*INDEX(\'Emission Factors\'!B:B,MATCH(IF(TRIM(J{r})="","Default road vehicle",TRIM(J{r})),\'Emission Factors\'!A:A,0)))/1000000,0)'
    if mode == 'rail':
        return f'=IF(F{r}="Rail",((C{r}/2204.62*M{r})*INDEX(\'Emission Factors\'!B:B,MATCH(IF(TRIM(H{r})="","Rail industry-average",TRIM(H{r})),\'Emission Factors\'!A:A,0)))/1000000,0)'
    if mode == 'inland':
        return f'=IF(F{r}="Inland Water",((C{r}/2204.62*M{r})*INDEX(\'Emission Factors\'!B:B,MATCH(IF(TRIM(L{r})="","Default inland water vehicle",TRIM(L{r})),\'Emission Factors\'!A:A,0)))/1000000,0)'
    if mode == 'sea':
        return f'=IF(F{r}="Sea",((C{r}/2204.62*M{r})*INDEX(\'Emission Factors\'!B:B,MATCH(IF(TRIM(K{r})="","Default sea vessel",TRIM(K{r})),\'Emission Factors\'!A:A,0)))/1000000,0)'
    if mode == 'storage':
        return f'=IF(F{r}="Storage",(C{r}/2204.62)*\'Emission Factors\'!B50/1000,0)'
    raise ValueError(mode)


def clear_row_a_to_n(ws, row_num):
    for c in COLS_A_TO_N:
        ws.cell(row_num, c).value = None


def process_workbooks(input_path, template_path, output_path=None, repair_formulas=True):
    input_wb = load_workbook(input_path, data_only=False)
    input_ws = input_wb[input_wb.sheetnames[0]]
    template_wb = load_workbook(template_path, data_only=False)
    output_ws = template_wb['Project Data and GHG Calcs']

    input_map = parse_headers(input_ws)
    missing_headers = [h for h in INPUT_HEADERS if h not in input_map]
    if missing_headers:
        raise ValueError(f'Input sheet is missing required headers: {", ".join(missing_headers)}')

    rows = []
    for row_num in range(2, input_ws.max_row + 1):
        row = {h: input_ws.cell(row_num, input_map[h]).value for h in INPUT_HEADERS}
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row.values()):
            continue
        weight = row['Weight (Pounds)']
        if isinstance(weight, str) and weight.strip().lower() == 'legal':
            row['Weight (Pounds)'] = 44000
        else:
            row['Weight (Pounds)'] = clean_output_value(weight)
        for key in INPUT_HEADERS:
            if key != 'Weight (Pounds)':
                row[key] = clean_output_value(row[key])
        if looks_blank_or_zero(row['Distance (Miles)']):
            estimate, method = estimate_distance(row['Origin'], row['Destination'], row['Mode'])
            row['Distance (Miles)'] = estimate
            row['Distance Estimation Method'] = method
        else:
            row['Distance Estimation Method'] = None
        rows.append(row)

    def sort_key(item):
        value = item['Delivery Date']
        if isinstance(value, datetime):
            return value
        return datetime.max

    rows.sort(key=sort_key)

    start_row = 2
    for idx, row in enumerate(rows, start=start_row):
        output_ws.cell(idx, 1).value = row['Project ID']
        output_ws.cell(idx, 2).value = row['Pick Up Date']
        output_ws.cell(idx, 3).value = row['Weight (Pounds)']
        output_ws.cell(idx, 4).value = row['Origin']
        output_ws.cell(idx, 5).value = row['Destination']
        output_ws.cell(idx, 6).value = row['Mode']
        output_ws.cell(idx, 7).value = row['Delivery Date']
        output_ws.cell(idx, 8).value = row['Subcontractor/Partner']
        output_ws.cell(idx, 9).value = row['Client']
        output_ws.cell(idx, 10).value = row['(Road) Vehicle category (US)']
        output_ws.cell(idx, 11).value = row['(Sea) Vessel Type']
        output_ws.cell(idx, 12).value = row['(Inland Water) Vessel Type']
        output_ws.cell(idx, 13).value = row['Distance (Miles)']
        output_ws.cell(idx, 14).value = row['Distance Estimation Method']
        if repair_formulas:
            output_ws.cell(idx, 15).value = compatible_formula(idx, 'road')
            output_ws.cell(idx, 16).value = compatible_formula(idx, 'rail')
            output_ws.cell(idx, 17).value = compatible_formula(idx, 'inland')
            output_ws.cell(idx, 18).value = compatible_formula(idx, 'sea')
            output_ws.cell(idx, 19).value = compatible_formula(idx, 'storage')

    for row_num in range(start_row + len(rows), output_ws.max_row + 1):
        clear_row_a_to_n(output_ws, row_num)

    if output_path is None:
        stamp = datetime.now().strftime('%Y-%m-%d')
        output_path = str(Path(template_path).with_name(f'HLI Project GHG Calcs {stamp}.xlsx'))

    template_wb.save(output_path)
    return output_path, len(rows)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry('760x540')
        self.input_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.repair_formulas = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value='Select the input workbook and the template workbook.')
        self._build()

    def _build(self):
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text='HLI Project GHG Calcs Builder', font=('Segoe UI', 16, 'bold')).pack(anchor='w')
        ttk.Label(
            frame,
            text='Processes workbooks locally. No data is stored in any external database. Only Origin, Destination, and Mode may be sent to online map services when a distance estimate is needed.',
            wraplength=700,
        ).pack(anchor='w', pady=(6, 14))

        self._file_row(frame, 'Input workbook', self.input_path, self.pick_input)
        self._file_row(frame, 'Template workbook', self.template_path, self.pick_template)
        self._file_row(frame, 'Output folder', self.output_dir, self.pick_output_dir)

        ttk.Checkbutton(
            frame,
            text='Repair template formulas in O:S for Excel compatibility (recommended)',
            variable=self.repair_formulas,
        ).pack(anchor='w', pady=(10, 8))

        info = (
            'Rules applied:\n'
            '• Writes only columns A:N\n'
            '• Leaves the Emission Factors tab intact\n'
            '• Converts Legal weight to 44,000\n'
            '• Leaves 0, NAME?, and #NAME? blank in imported values\n'
            '• Sorts by Delivery Date, oldest first\n'
            '• Estimates missing distances online\n'
            '• Rail distance = 90% of road distance\n'
        )
        ttk.Label(frame, text=info, justify='left').pack(anchor='w', pady=(6, 10))

        btns = ttk.Frame(frame)
        btns.pack(fill='x', pady=(8, 10))
        self.process_btn = ttk.Button(btns, text='Process Workbooks', command=self.start_processing)
        self.process_btn.pack(side='left')
        ttk.Button(btns, text='Open Output Folder', command=self.open_output_folder).pack(side='left', padx=8)

        ttk.Label(frame, textvariable=self.status, wraplength=700).pack(anchor='w', pady=(10, 0))

        sec = ttk.LabelFrame(frame, text='Data security statement', padding=12)
        sec.pack(fill='both', expand=True, pady=(18, 0))
        text = tk.Text(sec, height=9, wrap='word')
        text.pack(fill='both', expand=True)
        security_text = (
            "This application is designed to protect sensitive client data by performing all spreadsheet processing locally on the user’s device. "
            "Input files are not uploaded, stored, or transmitted to any external servers or databases.\n\n"
            "To support distance estimation where required, the application makes limited external requests to third-party routing and geocoding services. "
            "These requests include only non-sensitive geographic information (i.e., origin and destination locations) necessary to calculate transportation distances. "
            "No client-specific data, emissions data, financial information, or full spreadsheet contents are transmitted externally.\n\n"
            "The application does not retain, log, or store any data beyond the local session. "
            "All outputs are generated and saved directly on the user’s device."
        )
        text.insert('1.0', security_text)
        text.configure(state='disabled')

    def _file_row(self, parent, label, var, command):
        row = ttk.Frame(parent)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text=label, width=16).pack(side='left')
        ttk.Entry(row, textvariable=var).pack(side='left', fill='x', expand=True, padx=(0, 8))
        ttk.Button(row, text='Browse', command=command).pack(side='left')

    def pick_input(self):
        path = filedialog.askopenfilename(filetypes=[('Excel workbooks', '*.xlsx')])
        if path:
            self.input_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(str(Path(path).parent))

    def pick_template(self):
        path = filedialog.askopenfilename(filetypes=[('Excel workbooks', '*.xlsx')])
        if path:
            self.template_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(str(Path(path).parent))

    def pick_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def open_output_folder(self):
        path = self.output_dir.get().strip()
        if not path:
            messagebox.showinfo(APP_TITLE, 'Choose an output folder first.')
            return
        try:
            import os
            import subprocess
            import sys
            if sys.platform.startswith('win'):
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f'Could not open folder.\n\n{exc}')

    def start_processing(self):
        if not self.input_path.get().strip() or not self.template_path.get().strip():
            messagebox.showinfo(APP_TITLE, 'Please select both the input workbook and the template workbook.')
            return
        self.process_btn.configure(state='disabled')
        self.status.set('Processing... this may take a moment if online distance estimates are needed.')
        thread = threading.Thread(target=self._run_processing, daemon=True)
        thread.start()

    def _run_processing(self):
        try:
            output_dir = Path(self.output_dir.get().strip() or Path(self.template_path.get()).parent)
            output_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime('%Y-%m-%d')
            output_path = output_dir / f'HLI Project GHG Calcs {stamp}.xlsx'
            saved_path, row_count = process_workbooks(
                self.input_path.get().strip(),
                self.template_path.get().strip(),
                str(output_path),
                repair_formulas=self.repair_formulas.get(),
            )
            self.root.after(0, lambda: self._finish_success(saved_path, row_count))
        except Exception as exc:
            self.root.after(0, lambda: self._finish_error(exc))

    def _finish_success(self, saved_path, row_count):
        self.process_btn.configure(state='normal')
        self.status.set(f'Success. Saved {row_count} populated rows to {saved_path}')
        messagebox.showinfo(APP_TITLE, f'Success. Saved {row_count} populated rows.\n\n{saved_path}')

    def _finish_error(self, exc):
        self.process_btn.configure(state='normal')
        self.status.set(f'Processing failed: {exc}')
        messagebox.showerror(APP_TITLE, f'Processing failed.\n\n{exc}')


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use('clam')
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
