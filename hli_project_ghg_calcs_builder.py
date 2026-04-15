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
USER_AGENT = "HLI-Project-GHG-Calcs/2.0 (local desktop app)"
REQUEST_TIMEOUT = 12

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
    'los angeles, california': 'Los Angeles, CA',
    'cartagena, colombia': 'Cartagena, Colombia',
}

PORT_HINTS = {
    'the hague, netherlands': 'Port of Rotterdam, Netherlands',
    'new york, ny': 'Port of New York, NY',
    'new york city, ny': 'Port of New York, NY',
    'los angeles, california': 'Port of Los Angeles, CA',
    'los angeles, ca': 'Port of Los Angeles, CA',
    'cartagena, colombia': 'Port of Cartagena, Colombia',
    'port of freeport': 'Port Freeport, Freeport, TX',
    'port of houston': 'Port of Houston, Houston, TX',
    'myrtle beach, sc': 'Port of Charleston, SC',
}

CITY_COORDS = {
    'laredo, tx': (27.5306, -99.4803),
    'laredo, texas': (27.5306, -99.4803),
    'los angeles, ca': (34.0522, -118.2437),
    'los angeles, california': (34.0522, -118.2437),
    'cartagena, colombia': (10.3910, -75.4794),
    'new york, ny': (40.7128, -74.0060),
    'new york city, ny': (40.7128, -74.0060),
    'houston, tx': (29.7604, -95.3698),
    'freeport, tx': (28.9541, -95.3597),
    'charleston, sc': (32.7765, -79.9311),
    'rotterdam, netherlands': (51.9244, 4.4777),
    'the hague, netherlands': (52.0705, 4.3007),
    'port of los angeles, ca': (33.7361, -118.2626),
    'port of cartagena, colombia': (10.3997, -75.5144),
    'port of houston, houston, tx': (29.7284, -95.2650),
    'port freeport, freeport, tx': (28.9365, -95.3088),
    'port of rotterdam, netherlands': (51.8850, 4.2867),
    'port of new york, ny': (40.6840, -74.0419),
}

STATE_COORDS = {
    'al': (32.8067, -86.7911), 'alabama': (32.8067, -86.7911),
    'ak': (61.3707, -152.4044), 'alaska': (61.3707, -152.4044),
    'az': (33.7298, -111.4312), 'arizona': (33.7298, -111.4312),
    'ar': (34.9697, -92.3731), 'arkansas': (34.9697, -92.3731),
    'ca': (36.1162, -119.6816), 'california': (36.1162, -119.6816),
    'co': (39.0598, -105.3111), 'colorado': (39.0598, -105.3111),
    'ct': (41.5978, -72.7554), 'connecticut': (41.5978, -72.7554),
    'de': (39.3185, -75.5071), 'delaware': (39.3185, -75.5071),
    'fl': (27.7663, -81.6868), 'florida': (27.7663, -81.6868),
    'ga': (33.0406, -83.6431), 'georgia': (33.0406, -83.6431),
    'hi': (21.0943, -157.4983), 'hawaii': (21.0943, -157.4983),
    'id': (44.2405, -114.4788), 'idaho': (44.2405, -114.4788),
    'il': (40.3495, -88.9861), 'illinois': (40.3495, -88.9861),
    'in': (39.8494, -86.2583), 'indiana': (39.8494, -86.2583),
    'ia': (42.0115, -93.2105), 'iowa': (42.0115, -93.2105),
    'ks': (38.5266, -96.7265), 'kansas': (38.5266, -96.7265),
    'ky': (37.6681, -84.6701), 'kentucky': (37.6681, -84.6701),
    'la': (31.1695, -91.8678), 'louisiana': (31.1695, -91.8678),
    'me': (44.6939, -69.3819), 'maine': (44.6939, -69.3819),
    'md': (39.0639, -76.8021), 'maryland': (39.0639, -76.8021),
    'ma': (42.2302, -71.5301), 'massachusetts': (42.2302, -71.5301),
    'mi': (43.3266, -84.5361), 'michigan': (43.3266, -84.5361),
    'mn': (45.6945, -93.9002), 'minnesota': (45.6945, -93.9002),
    'ms': (32.7416, -89.6787), 'mississippi': (32.7416, -89.6787),
    'mo': (38.4561, -92.2884), 'missouri': (38.4561, -92.2884),
    'mt': (46.9219, -110.4544), 'montana': (46.9219, -110.4544),
    'ne': (41.1254, -98.2681), 'nebraska': (41.1254, -98.2681),
    'nv': (38.3135, -117.0554), 'nevada': (38.3135, -117.0554),
    'nh': (43.4525, -71.5639), 'new hampshire': (43.4525, -71.5639),
    'nj': (40.2989, -74.5210), 'new jersey': (40.2989, -74.5210),
    'nm': (34.8405, -106.2485), 'new mexico': (34.8405, -106.2485),
    'ny': (42.1657, -74.9481), 'new york': (42.1657, -74.9481),
    'nc': (35.6301, -79.8064), 'north carolina': (35.6301, -79.8064),
    'nd': (47.5289, -99.7840), 'north dakota': (47.5289, -99.7840),
    'oh': (40.3888, -82.7649), 'ohio': (40.3888, -82.7649),
    'ok': (35.5653, -96.9289), 'oklahoma': (35.5653, -96.9289),
    'or': (44.5720, -122.0709), 'oregon': (44.5720, -122.0709),
    'pa': (40.5908, -77.2098), 'pennsylvania': (40.5908, -77.2098),
    'ri': (41.6809, -71.5118), 'rhode island': (41.6809, -71.5118),
    'sc': (33.8569, -80.9450), 'south carolina': (33.8569, -80.9450),
    'sd': (44.2998, -99.4388), 'south dakota': (44.2998, -99.4388),
    'tn': (35.7478, -86.6923), 'tennessee': (35.7478, -86.6923),
    'tx': (31.0545, -97.5635), 'texas': (31.0545, -97.5635),
    'ut': (40.1500, -111.8624), 'utah': (40.1500, -111.8624),
    'vt': (44.0459, -72.7107), 'vermont': (44.0459, -72.7107),
    'va': (37.7693, -78.1700), 'virginia': (37.7693, -78.1700),
    'wa': (47.4009, -121.4905), 'washington': (47.4009, -121.4905),
    'wv': (38.4912, -80.9545), 'west virginia': (38.4912, -80.9545),
    'wi': (44.2685, -89.6165), 'wisconsin': (44.2685, -89.6165),
    'wy': (42.7560, -107.3025), 'wyoming': (42.7560, -107.3025),
}

COUNTRY_COORDS = {
    'united states': (39.8283, -98.5795), 'usa': (39.8283, -98.5795), 'us': (39.8283, -98.5795),
    'mexico': (23.6345, -102.5528),
    'canada': (56.1304, -106.3468),
    'colombia': (4.5709, -74.2973),
    'netherlands': (52.1326, 5.2913),
    'saudi arabia': (23.8859, 45.0792),
    'china': (35.8617, 104.1954),
    'japan': (36.2048, 138.2529),
    'south korea': (35.9078, 127.7669),
    'india': (20.5937, 78.9629),
    'germany': (51.1657, 10.4515),
    'france': (46.2276, 2.2137),
    'spain': (40.4637, -3.7492),
    'italy': (41.8719, 12.5674),
    'united kingdom': (55.3781, -3.4360), 'uk': (55.3781, -3.4360),
    'brazil': (-14.2350, -51.9253),
    'argentina': (-38.4161, -63.6167),
    'chile': (-35.6751, -71.5430),
    'peru': (-9.1900, -75.0152),
    'australia': (-25.2744, 133.7751),
    'new zealand': (-40.9006, 174.8860),
    'singapore': (1.3521, 103.8198),
    'united arab emirates': (23.4241, 53.8478), 'uae': (23.4241, 53.8478),
}

SAME_REGION_DEFAULTS = {
    'road': 150,
    'rail': 180,
    'sea': 250,
    'inland water': 80,
}

MODE_FACTORS = {
    'road_gc': 1.18,
    'rail_gc': 1.22,
    'sea_gc': 1.30,
    'inland_gc': 1.15,
}

_geocode_cache = {}
_route_cache = {}
_distance_cache = {}


def normalize_text(value):
    if value is None:
        return ''
    text = str(value).strip()
    lowered = re.sub(r'\s+', ' ', text.lower())
    return ALIASES.get(lowered, text)


def normalize_key(value):
    return re.sub(r'\s+', ' ', normalize_text(value).strip().lower())


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


def _make_geo_result(lat, lon, label, source):
    return {'lat': float(lat), 'lon': float(lon), 'display_name': label, 'source': source}


def _candidate_texts(query):
    q = normalize_text(query)
    candidates = []
    if q:
        candidates.append(q)
    low = normalize_key(q)
    if low in PORT_HINTS:
        candidates.append(PORT_HINTS[low])
    compact = q.replace('Port of ', '').replace('Port ', '')
    if compact and compact not in candidates:
        candidates.append(compact)
    for cand in list(candidates):
        parts = [p.strip() for p in cand.split(',') if p.strip()]
        if len(parts) >= 2:
            simplified = ', '.join(parts[:2])
            if simplified not in candidates:
                candidates.append(simplified)
    return candidates


def _lookup_builtin(query):
    for cand in _candidate_texts(query):
        key = normalize_key(cand)
        if key in CITY_COORDS:
            lat, lon = CITY_COORDS[key]
            return _make_geo_result(lat, lon, cand, 'builtin-city')
        parts = [p.strip().lower() for p in cand.split(',') if p.strip()]
        for part in reversed(parts):
            if part in STATE_COORDS:
                lat, lon = STATE_COORDS[part]
                return _make_geo_result(lat, lon, cand, 'builtin-state')
            if part in COUNTRY_COORDS:
                lat, lon = COUNTRY_COORDS[part]
                return _make_geo_result(lat, lon, cand, 'builtin-country')
    return None


def geocode_location(query):
    query = normalize_text(query)
    if not query:
        return None
    if query in _geocode_cache:
        return _geocode_cache[query]

    builtin = _lookup_builtin(query)
    if builtin:
        _geocode_cache[query] = builtin
        return builtin

    for candidate in _candidate_texts(query):
        try:
            response = requests.get(
                'https://nominatim.openstreetmap.org/search',
                params={'format': 'jsonv2', 'limit': 1, 'q': candidate},
                headers={'User-Agent': USER_AGENT},
                timeout=REQUEST_TIMEOUT,
            )
            response.raise_for_status()
            data = response.json()
            if data:
                result = {
                    'lat': float(data[0]['lat']),
                    'lon': float(data[0]['lon']),
                    'display_name': data[0]['display_name'],
                    'source': 'online-geocode',
                }
                _geocode_cache[query] = result
                return result
        except Exception:
            continue

    builtin = _lookup_builtin(query)
    _geocode_cache[query] = builtin
    return builtin


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


def _parts(text):
    return [p.strip().lower() for p in normalize_text(text).split(',') if p.strip()]


def _same_region(origin, destination):
    op = _parts(origin)
    dp = _parts(destination)
    if normalize_key(origin) == normalize_key(destination):
        return 'same-location'
    if op and dp and op[-1] == dp[-1]:
        if len(op) > 1 and len(dp) > 1 and op[-2] == dp[-2]:
            return 'same-subregion'
        return 'same-country-or-state'
    return 'different-region'


def fallback_distance_miles(origin, destination, mode):
    region = _same_region(origin, destination)
    mode_key = (mode or '').strip().lower()
    if region == 'same-location':
        miles = SAME_REGION_DEFAULTS.get(mode_key, 100)
    elif region == 'same-subregion':
        miles = {'road': 225, 'rail': 260, 'sea': 400, 'inland water': 140}.get(mode_key, 250)
    elif region == 'same-country-or-state':
        miles = {'road': 450, 'rail': 600, 'sea': 800, 'inland water': 250}.get(mode_key, 500)
    else:
        miles = {'road': 900, 'rail': 1400, 'sea': 3500, 'inland water': 600}.get(mode_key, 1200)
    method = f'Estimated using deterministic {mode_key or "transport"} fallback heuristic'
    return miles, method


def sea_or_water_distance_miles(origin, destination, mode):
    mode_key = (mode or '').strip().lower()
    origin_q = PORT_HINTS.get(normalize_key(origin), normalize_text(origin))
    destination_q = PORT_HINTS.get(normalize_key(destination), normalize_text(destination))
    a = geocode_location(origin_q)
    b = geocode_location(destination_q)
    if a and b:
        base = haversine_miles(a, b)
        factor = MODE_FACTORS['sea_gc'] if mode_key == 'sea' else MODE_FACTORS['inland_gc']
        miles = max(1, round(base * factor))
        label = 'Sea estimate' if mode_key == 'sea' else 'Inland water estimate'
        method = f'{label} (great-circle × calibrated factor)'
        if origin_q != normalize_text(origin) or destination_q != normalize_text(destination):
            method += f' using port proxies: {origin_q} to {destination_q}'
        return miles, method
    return fallback_distance_miles(origin, destination, mode)


def estimate_distance(origin, destination, mode):
    mode_key = (mode or '').strip().lower()
    cache_key = (normalize_key(origin), normalize_key(destination), mode_key)
    if cache_key in _distance_cache:
        return _distance_cache[cache_key]

    if mode_key == 'storage':
        result = (None, None)
    elif mode_key == 'road':
        road = road_distance_miles(origin, destination)
        if road is not None:
            result = (max(1, round(road)), 'Road estimate (online network routing)')
        else:
            a = geocode_location(origin)
            b = geocode_location(destination)
            if a and b:
                miles = max(1, round(haversine_miles(a, b) * MODE_FACTORS['road_gc']))
                result = (miles, 'Road estimate (great-circle × calibrated factor)')
            else:
                result = fallback_distance_miles(origin, destination, mode)
    elif mode_key == 'rail':
        a = geocode_location(origin)
        b = geocode_location(destination)
        if a and b:
            gc = haversine_miles(a, b)
            miles = max(1, round(gc * MODE_FACTORS['rail_gc']))
            result = (miles, 'Rail estimate (great-circle × calibrated rail factor)')
        else:
            result = fallback_distance_miles(origin, destination, mode)
    elif mode_key in {'sea', 'inland water'}:
        result = sea_or_water_distance_miles(origin, destination, mode)
    else:
        result = fallback_distance_miles(origin, destination, mode)

    _distance_cache[cache_key] = result
    return result


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
        self.root.geometry('780x560')
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
            text='Processes workbooks locally. No data is stored in any external database. When needed, limited online map calls may be used for distance estimation, with deterministic local fallbacks so non-Storage distances are always populated.',
            wraplength=720,
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
            '• Distance never remains blank except for Storage\n'
            '• Reuses cached lane estimates for faster processing\n'
            '• Rail uses great-circle × calibrated rail factor as the primary estimate\n'
        )
        ttk.Label(frame, text=info, justify='left').pack(anchor='w', pady=(6, 10))

        btns = ttk.Frame(frame)
        btns.pack(fill='x', pady=(8, 10))
        self.process_btn = ttk.Button(btns, text='Process Workbooks', command=self.start_processing)
        self.process_btn.pack(side='left')
        ttk.Button(btns, text='Open Output Folder', command=self.open_output_folder).pack(side='left', padx=8)

        ttk.Label(frame, textvariable=self.status, wraplength=720).pack(anchor='w', pady=(10, 0))

        sec = ttk.LabelFrame(frame, text='Data security statement', padding=12)
        sec.pack(fill='both', expand=True, pady=(18, 0))
        text = tk.Text(sec, height=9, wrap='word')
        text.pack(fill='both', expand=True)
        security_text = (
            "This application is designed to protect sensitive client data by performing all spreadsheet processing locally on the user’s device. "
            "Input files are not uploaded, stored, or transmitted to any external servers or databases.\n\n"
            "To support distance estimation where required, the application may make limited external requests to third-party routing and geocoding services. "
            "These requests include only non-sensitive geographic information (i.e., origin and destination locations) necessary to calculate transportation distances. "
            "No client-specific data, emissions data, financial information, or full spreadsheet contents are transmitted externally.\n\n"
            "When online routing or geocoding is unavailable, the application falls back to deterministic local estimation methods so that all non-Storage shipments still receive a distance estimate. "
            "The application does not retain, log, or store any data beyond the local session. All outputs are generated and saved directly on the user’s device."
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
        self.status.set('Processing... lane estimates are cached for speed.')
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
