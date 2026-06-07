#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aplikacja do weryfikacji faktur Żabka
"""

import os
import sys
import json
import socket
import webbrowser
import threading
import shutil
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET
from flask import Flask, jsonify, request, send_file, render_template_string, abort

# ─── ŚCIEŻKI ──────────────────────────────────────────────────────────────────

def _pick_folder():
    """Pokazuje systemowy dialog wyboru folderu. Zwraca wybraną ścieżkę lub None."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder = filedialog.askdirectory(
            title='Wybierz folder z fakturami',
            initialdir=_load_last_folder(),
        )
        root.destroy()
        return folder or None
    except Exception as e:
        print(f'[BŁĄD] Dialog wyboru folderu: {e}', file=sys.stderr)
        return None

def _config_path():
    """Ścieżka do pliku konfiguracyjnego w AppData (Windows) lub ~/.config (inne)."""
    appdata = os.environ.get('APPDATA')  # Windows: C:\Users\...\AppData\Roaming
    if appdata:
        config_dir = os.path.join(appdata, 'FakturyZabka')
    else:
        config_dir = os.path.join(os.path.expanduser('~'), '.faktury_zabka')
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, 'config.json')

def _load_last_folder():
    try:
        with open(_config_path(), encoding='utf-8') as f:
            return json.load(f).get('last_folder', os.path.expanduser('~'))
    except Exception:
        return os.path.expanduser('~')

def _save_last_folder(path):
    try:
        with open(_config_path(), 'w', encoding='utf-8') as f:
            json.dump({'last_folder': path}, f, ensure_ascii=False)
    except Exception:
        pass

# Ustal BASE_DIR
if '--folder' in sys.argv:
    # Tryb deweloperski: python app.py --folder /ścieżka
    idx = sys.argv.index('--folder')
    BASE_DIR = os.path.abspath(sys.argv[idx + 1])
elif getattr(sys, 'frozen', False):
    # Uruchomiony jako .exe — pokaż dialog wyboru folderu
    chosen = _pick_folder()
    if not chosen:
        sys.exit(0)  # użytkownik anulował
    BASE_DIR = chosen
    _save_last_folder(BASE_DIR)
else:
    # Tryb deweloperski bez --folder: użyj folderu skryptu
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROGRESS_FILE = os.path.join(BASE_DIR, 'postep.json')
VERIFIED_BASE = os.path.join(BASE_DIR, 'zweryfikowane')

# ─── FORMATOWANIE ─────────────────────────────────────────────────────────────

def fmt_date(d):
    """'2026-03-30' → '30.03.2026'"""
    if not d or len(d) < 10:
        return d or '—'
    try:
        return f'{d[8:10]}.{d[5:7]}.{d[:4]}'
    except Exception:
        return d

def fmt_money(v):
    """'1366.09' → '1 366,09 PLN'"""
    if not v:
        return None
    try:
        num = float(v)
        integer, frac = f'{num:.2f}'.split('.')
        result = ''
        for i, ch in enumerate(reversed(integer)):
            if i and i % 3 == 0:
                result = ' ' + result
            result = ch + result
        return f'{result},{frac} PLN'
    except Exception:
        return str(v)

# ─── XML PARSOWANIE ───────────────────────────────────────────────────────────

DEFAULT_NS = 'http://crd.gov.pl/wzor/2025/06/25/13775/'

def detect_ns(xml_path):
    try:
        for event, elem in ET.iterparse(xml_path, events=['start-ns']):
            if elem[0] == '':
                return elem[1]
    except Exception:
        pass
    return DEFAULT_NS

def ft(root, ns, tag):
    """Znajdź pierwszy element i zwróć jego tekst."""
    el = root.find(f'.//{{{ns}}}{tag}')
    return el.text.strip() if el is not None and el.text else None

def fta(root, ns, tag):
    """Znajdź wszystkie elementy i zwróć listę tekstów."""
    return [el.text.strip() for el in root.findall(f'.//{{{ns}}}{tag}') if el.text]

def parse_wz(wz_list):
    """Wyciąga tylko numer WZ: '860315796 / 25.03.2026 ( 0511140/26/ACWR )' → '860315796'"""
    result = []
    for wz in wz_list:
        num = wz.split('/')[0].strip()
        if num:
            result.append(num)
    return result

def parse_invoice_file(xml_path):
    try:
        ns   = detect_ns(xml_path)
        root = ET.parse(xml_path).getroot()
        num  = ft(root, ns, 'P_2')
        if not num:
            return None
        if num.startswith('63'):
            return _p63(root, ns, num)
        elif num.startswith('66'):
            return _p66(root, ns, num)
        elif num.startswith('17'):
            return _p17(root, ns, num)
        elif num.startswith('15'):
            return _p15(root, ns, num)
        return _p_inne(root, ns, num)
    except Exception as e:
        print(f'[BŁĄD] Parsowanie {os.path.basename(xml_path)}: {e}', file=sys.stderr)
        return None

def _iso(root, ns):
    """Zwraca datę wystawienia w formacie ISO (YYYY-MM-DD) do sortowania/filtrowania."""
    raw = ft(root, ns, 'P_1') or ''
    return raw[:10] if len(raw) >= 10 else ''

def _collect_row_amounts(root, ns):
    """Zbiera wartości pozycji (P_11) ze wszystkich wierszy FaWiersz."""
    vals = []
    for row in root.findall(f'.//{{{ns}}}FaWiersz'):
        v = row.findtext(f'{{{ns}}}P_11')
        if v and v.strip():
            vals.append(v.strip())
    return vals

def _p63(root, ns, num):
    total_raw     = ft(root, ns, 'P_15')
    do_zaplaty_r  = ft(root, ns, 'DoZaplaty')
    suma_obc_r    = ft(root, ns, 'SumaObciazen')
    row_amounts   = _collect_row_amounts(root, ns)
    amounts_raw   = [a for a in [total_raw, do_zaplaty_r, suma_obc_r] + row_amounts if a]
    return {
        'type':           '63',
        'invoice_num':    num,
        'issue_date':     fmt_date(ft(root, ns, 'P_1')),
        'issue_date_iso': _iso(root, ns),
        'service_date':   fmt_date(ft(root, ns, 'P_6')),
        'wz_numbers':     parse_wz(fta(root, ns, 'WZ')),
        'total':          fmt_money(total_raw),
        'total_raw':      total_raw,
        'do_zaplaty':     fmt_money(do_zaplaty_r),
        'suma_obciazen':  fmt_money(suma_obc_r),
        'amounts_raw':    amounts_raw,
    }

def _p66(root, ns, num):
    total_raw   = ft(root, ns, 'P_15')
    row_amounts = _collect_row_amounts(root, ns)
    amounts_raw = [a for a in [total_raw] + row_amounts if a]
    return {
        'type':           '66',
        'invoice_num':    num,
        'issue_date':     fmt_date(ft(root, ns, 'P_1')),
        'issue_date_iso': _iso(root, ns),
        'service_date':   fmt_date(ft(root, ns, 'P_6')),
        'wz_numbers':     parse_wz(fta(root, ns, 'WZ')),
        'total':          fmt_money(total_raw),
        'total_raw':      total_raw,
        'amounts_raw':    amounts_raw,
    }

def _p15(root, ns, num):
    """Farmacja — podobna do 63, z WZ i kwotą do zapłaty."""
    total_raw    = ft(root, ns, 'P_15')
    do_zaplaty_r = ft(root, ns, 'DoZaplaty')
    row_amounts  = _collect_row_amounts(root, ns)
    amounts_raw  = [a for a in [total_raw, do_zaplaty_r] + row_amounts if a]
    return {
        'type':           '15',
        'invoice_num':    num,
        'issue_date':     fmt_date(ft(root, ns, 'P_1')),
        'issue_date_iso': _iso(root, ns),
        'service_date':   fmt_date(ft(root, ns, 'P_6')),
        'wz_numbers':     parse_wz(fta(root, ns, 'WZ')),
        'total':          fmt_money(total_raw),
        'total_raw':      total_raw,
        'do_zaplaty':     fmt_money(do_zaplaty_r),
        'amounts_raw':    amounts_raw,
    }

def _p_billbird(root, ns, num):
    """Faktury 17 z 'Prowizja od rachunków BillBird' — bez EVOUCHERów, tylko DoZaplaty."""
    total_raw    = ft(root, ns, 'P_15')
    do_zaplaty_r = ft(root, ns, 'DoZaplaty')
    amounts_raw  = [a for a in [total_raw, do_zaplaty_r] if a]
    return {
        'type':           'billbird',
        'invoice_num':    num,
        'issue_date':     fmt_date(ft(root, ns, 'P_1')),
        'issue_date_iso': _iso(root, ns),
        'service_date':   fmt_date(ft(root, ns, 'P_6')),
        'total':          fmt_money(total_raw),
        'total_raw':      total_raw,
        'do_zaplaty':     fmt_money(do_zaplaty_r),
        'do_zaplaty_raw': do_zaplaty_r,
        'amounts_raw':    amounts_raw,
    }

def _p_inne(root, ns, num):
    total_raw    = ft(root, ns, 'P_15')
    korygowana   = ft(root, ns, 'NrFaKorygowanej') or ft(root, ns, 'NumerFakturyKorygowanej')
    return {
        'type':                      'inne',
        'invoice_num':               num,
        'issue_date':                fmt_date(ft(root, ns, 'P_1')),
        'issue_date_iso':            _iso(root, ns),
        'service_date':              fmt_date(ft(root, ns, 'P_6')),
        'total':                     fmt_money(total_raw),
        'total_raw':                 total_raw,
        'amounts_raw':               [total_raw] if total_raw else [],
        'numer_faktury_korygowanej': korygowana,
    }

def _p17(root, ns, num):
    totals = defaultdict(lambda: {'qty': 0.0, 'val': 0.0, 'price': 0.0})

    for row in root.findall(f'.//{{{ns}}}FaWiersz'):
        name_el = row.find(f'{{{ns}}}P_7')
        if name_el is None or not name_el.text:
            continue
        name = name_el.text.strip()
        if 'EVOUCHER' not in name.upper():
            continue
        qty   = float(row.findtext(f'{{{ns}}}P_8B', '0') or '0')
        val   = float(row.findtext(f'{{{ns}}}P_11',  '0') or '0')
        price = float(row.findtext(f'{{{ns}}}P_9A',  '0') or '0')
        totals[name]['qty'] += qty
        totals[name]['val'] += val
        if not totals[name]['price']:
            totals[name]['price'] = price

    if not totals:
        # Sprawdź czy to faktura BillBird (Prowizja od rachunków BillBird)
        for row in root.findall(f'.//{{{ns}}}FaWiersz'):
            p7 = row.findtext(f'{{{ns}}}P_7', '') or ''
            if 'BILLBIRD' in p7.upper():
                return _p_billbird(root, ns, num)
        # Żadne EVOUCHERy ani BillBird – traktuj jak "Inne"
        return _p_inne(root, ns, num)

    svc = fmt_date(ft(root, ns, 'P_6') or ft(root, ns, 'DataZamowienia'))

    total_raw = ft(root, ns, 'P_15')
    issue_iso = _iso(root, ns)
    items = []
    amounts_raw = [total_raw] if total_raw else []
    for name in sorted(totals.keys()):
        d = totals[name]
        q = d['qty']
        val_s   = str(round(d['val'], 2))
        price_s = str(round(d['price'], 2))
        items.append({
            'name':       name,
            'qty':        int(q) if q == int(q) else q,
            'value':      fmt_money(val_s),
            'unit_price': fmt_money(price_s),
        })
        amounts_raw.extend([val_s, price_s])

    return {
        'type':           '17',
        'invoice_num':    num,
        'issue_date':     fmt_date(ft(root, ns, 'P_1')),
        'issue_date_iso': issue_iso,
        'service_date':   svc,
        'evoucher_items': items,
        'total':          fmt_money(total_raw),
        'total_raw':      total_raw,
        'amounts_raw':    amounts_raw,
    }

# ─── WYPAKOWYWANIE ZIPÓW ──────────────────────────────────────────────────────

ZIPS_ARCHIVE = '_paczki'  # podfolder na już wypakowane zipa

def extract_zips(root_dir):
    """Szuka plików .zip w root_dir, wypakowuje je i przenosi do _paczki/.
    Zwraca listę wypakowanych nazw plików."""
    extracted = []
    try:
        entries = os.listdir(root_dir)
    except Exception:
        return extracted

    zips = [f for f in entries if f.lower().endswith('.zip')]
    if not zips:
        return extracted

    archive_dir = os.path.join(root_dir, ZIPS_ARCHIVE)
    os.makedirs(archive_dir, exist_ok=True)

    for fn in zips:
        zip_path = os.path.join(root_dir, fn)
        try:
            with zipfile.ZipFile(zip_path, 'r') as zf:
                zf.extractall(root_dir)
            dst = os.path.join(archive_dir, fn)
            # Jeśli plik o tej nazwie już istnieje w archiwum, dodaj sufiks
            if os.path.exists(dst):
                base, ext = os.path.splitext(fn)
                i = 1
                while os.path.exists(dst):
                    dst = os.path.join(archive_dir, f'{base}_{i}{ext}')
                    i += 1
            shutil.move(zip_path, dst)
            extracted.append(fn)
            print(f'[ZIP] Wypakowano: {fn}', file=sys.stderr)
        except Exception as e:
            print(f'[BŁĄD] Rozpakowywanie {fn}: {e}', file=sys.stderr)

    return extracted

# ─── PLIKI FAKTUR ─────────────────────────────────────────────────────────────

def iter_invoice_files(root_dir):
    """Rekurencyjnie przechodzi przez folder (pomijając 'zweryfikowane'),
    zwraca ścieżki do plików XML."""
    verified_abs = os.path.abspath(VERIFIED_BASE)
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # Pomiń folder zweryfikowanych i jego podfoldery
        if os.path.abspath(dirpath).startswith(verified_abs):
            dirnames.clear()
            continue
        # Pomiń ukryte foldery i folder z archiwum zipów
        dirnames[:] = [d for d in dirnames if not d.startswith('.') and d != ZIPS_ARCHIVE]
        for fn in sorted(filenames):
            if fn.lower().endswith('.xml'):
                yield os.path.join(dirpath, fn)

def find_files(num):
    """Zwraca (pdf_path, xml_path) dla numeru faktury — szuka rekurencyjnie.
    Zwraca pierwszy znaleziony plik każdego typu (dla kompatybilności z API)."""
    pdf, xml = None, None
    for fp in find_all_invoice_files(num, include_verified=True):
        lo = fp.lower()
        if lo.endswith('.pdf'):
            pdf = pdf or fp
        elif lo.endswith('.xml'):
            xml = xml or fp
    return pdf, xml

def find_all_invoice_files(num, include_verified=False):
    """Zwraca listę WSZYSTKICH plików (PDF i XML) powiązanych z numerem faktury."""
    found = []
    verified_abs = os.path.abspath(VERIFIED_BASE)

    # Szukaj w głównym folderze (rekurencyjnie, pomijając zweryfikowane)
    for dirpath, dirnames, filenames in os.walk(BASE_DIR):
        if os.path.abspath(dirpath).startswith(verified_abs):
            dirnames.clear()
            continue
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        for fn in filenames:
            if num not in fn:
                continue
            lo = fn.lower()
            if lo.endswith('.pdf') or lo.endswith('.xml'):
                found.append(os.path.join(dirpath, fn))

    # Opcjonalnie szukaj też w folderach zweryfikowanych
    if include_verified:
        for t in VERIFIED_SUBDIRS:
            vdir = os.path.join(VERIFIED_BASE, t)
            if not os.path.isdir(vdir):
                continue
            for fn in os.listdir(vdir):
                if num not in fn:
                    continue
                lo = fn.lower()
                if lo.endswith('.pdf') or lo.endswith('.xml'):
                    found.append(os.path.join(vdir, fn))

    return found

# ─── CACHE FAKTUR ─────────────────────────────────────────────────────────────

_cache = None
_skipped_nums = set()  # numery faktur z XML, które zostały celowo pominięte (np. typ 17 bez EVOUCHERów)

def num_from_pdf_filename(fn):
    """'6310390794 WIZU.pdf' → '6310390794'  (pierwsza część przed spacją)"""
    name = os.path.splitext(fn)[0]
    parts = name.split()
    return parts[0] if parts else None

def type_from_num(num):
    if num.startswith('63'): return '63'
    if num.startswith('66'): return '66'
    if num.startswith('17'): return '17'
    if num.startswith('15'): return '15'
    return 'inne'

VERIFIED_SUBDIRS = ('63', '66', '17', '15', 'BillBird', 'Inne')

def _fallback_pdf_entry(num):
    """Minimalny wpis dla faktury bez XML — tylko PDF."""
    return {
        'type':         type_from_num(num),
        'invoice_num':  num,
        'issue_date':   '—',
        'service_date': '—',
        'total':        None,
        'pdf_only':     True,   # brak XML
    }

def _scan_pdf_fallbacks(search_dir, inv, skip_verified=True):
    """Dodaje do inv faktury znane tylko z PDF (brak lub uszkodzony XML)."""
    verified_abs = os.path.abspath(VERIFIED_BASE)
    for dirpath, dirnames, filenames in os.walk(search_dir):
        if skip_verified and os.path.abspath(dirpath).startswith(verified_abs):
            dirnames.clear()
            continue
        dirnames[:] = [d for d in dirnames if not d.startswith('.') and d != ZIPS_ARCHIVE]
        for fn in filenames:
            if not fn.lower().endswith('.pdf'):
                continue
            num = num_from_pdf_filename(fn)
            if not num or num in inv or num in _skipped_nums:
                continue
            inv[num] = _fallback_pdf_entry(num)

def auto_verify_inne(inv, prog):
    """Automatycznie przenosi i zatwierdza wszystkie niezweryfikowane faktury typu 'inne'."""
    changed = False
    dest_dir = os.path.join(VERIFIED_BASE, 'Inne')

    for num, d in list(inv.items()):
        if d['type'] != 'inne':
            continue
        if prog.get(num, {}).get('verified', False):
            continue

        # Przenieś WSZYSTKIE pliki (może być kilka PDF/XML) do zweryfikowane/Inne/
        all_files = find_all_invoice_files(num, include_verified=False)
        os.makedirs(dest_dir, exist_ok=True)
        moved_paths = []
        for fpath in all_files:
            if os.path.exists(fpath):
                dst = os.path.join(dest_dir, os.path.basename(fpath))
                if not os.path.exists(dst):
                    shutil.move(fpath, dst)
                    moved_paths.append(fpath)
                    print(f'[INNE] Przeniesiono: {os.path.basename(fpath)}', file=sys.stderr)

        # Usuń puste podfoldery po przeniesieniu
        for fpath in moved_paths:
            parent = os.path.dirname(fpath)
            if parent != BASE_DIR:
                try:
                    if not os.listdir(parent):
                        os.rmdir(parent)
                except Exception:
                    pass

        # Oznacz jako zweryfikowane
        if num not in prog:
            prog[num] = {}
        prog[num]['verified'] = True
        changed = True

    if changed:
        save_progress(prog)

    return changed


def load_invoices(force=False):
    global _cache, _skipped_nums
    if _cache is not None and not force:
        return _cache

    # Wyczyść pominięte numery przy pełnym przeładowaniu
    _skipped_nums = set()

    # Wypakuj ewentualne nowe zipie przed skanowaniem
    extract_zips(BASE_DIR)

    inv = {}

    # Skanuj główny folder rekurencyjnie (XML)
    for xml_path in iter_invoice_files(BASE_DIR):
        d = parse_invoice_file(xml_path)
        if d:
            inv[d['invoice_num']] = d

    # Fallback: PDF bez XML w głównym folderze
    _scan_pdf_fallbacks(BASE_DIR, inv)

    # Foldery zweryfikowanych — najpierw XML, potem PDF (żeby XML miał zawsze priorytet)
    for t in VERIFIED_SUBDIRS:
        vdir = os.path.join(VERIFIED_BASE, t)
        if not os.path.isdir(vdir):
            continue
        fns = sorted(os.listdir(vdir))
        for fn in fns:
            if fn.lower().endswith('.xml'):
                d = parse_invoice_file(os.path.join(vdir, fn))
                if d and d['invoice_num'] not in inv:
                    inv[d['invoice_num']] = d
        for fn in fns:
            if fn.lower().endswith('.pdf'):
                num = num_from_pdf_filename(fn)
                if num and num not in inv:
                    inv[num] = _fallback_pdf_entry(num)

    # Auto-weryfikacja faktur "Inne" — przenosi pliki i oznacza jako sprawdzone
    prog = load_progress()
    auto_verify_inne(inv, prog)

    _cache = inv
    return inv

# ─── POSTĘP ───────────────────────────────────────────────────────────────────

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_progress(p):
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(p, f, ensure_ascii=False, indent=2)

# ─── FLASK ────────────────────────────────────────────────────────────────────

app = Flask(__name__)

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/api/invoices')
def api_invoices():
    inv  = load_invoices()
    prog = load_progress()
    # Odwrotny indeks: które faktury mają wystawioną korektę
    ma_korekte = {
        d['numer_faktury_korygowanej']
        for d in inv.values()
        if d.get('numer_faktury_korygowanej')
    }
    out  = []
    for num, d in sorted(inv.items()):
        p = prog.get(num, {})
        out.append({
            'invoice_num':    num,
            'type':           d['type'],
            'issue_date':     d.get('issue_date', '—'),
            'issue_date_iso': d.get('issue_date_iso', ''),
            'total':          d.get('total', '—'),
            'total_raw':      d.get('total_raw'),
            'verified':       p.get('verified', False),
            'starred':        p.get('starred', False),
            'wz_numbers':     d.get('wz_numbers', []),
            'amounts_raw':              d.get('amounts_raw', []),
            'numer_faktury_korygowanej': d.get('numer_faktury_korygowanej'),
            'has_korekta':    num in ma_korekte,
        })
    return jsonify(out)

@app.route('/api/invoice/<num>')
def api_invoice(num):
    inv = load_invoices()
    if num not in inv:
        abort(404)
    d    = dict(inv[num])
    prog = load_progress()
    inv_prog = prog.get(num, {})
    d['verified']   = inv_prog.get('verified', False)
    d['starred']    = inv_prog.get('starred', False)
    d['note']       = inv_prog.get('note', '')
    d['wz_checked'] = inv_prog.get('wz_checked', {})
    d['ev_checked'] = inv_prog.get('ev_checked', {})
    d['pdf_only']   = d.get('pdf_only', False)
    pdf, _ = find_files(num)
    d['pdf_filename'] = os.path.basename(pdf) if pdf else None
    return jsonify(d)

@app.route('/api/verify/<num>', methods=['POST'])
def api_verify(num):
    global _cache
    inv = load_invoices()
    if num not in inv:
        abort(404)

    inv_type  = inv[num]['type']
    if inv_type == 'inne':        dest_name = 'Inne'
    elif inv_type == 'billbird':  dest_name = 'BillBird'
    else:                         dest_name = inv_type
    dest = os.path.join(VERIFIED_BASE, dest_name)
    os.makedirs(dest, exist_ok=True)

    # Znajdź WSZYSTKIE pliki faktury rekurencyjnie (może być kilka PDF/XML)
    files_to_move = find_all_invoice_files(num, include_verified=False)

    moved = []
    for src in files_to_move:
        fn = os.path.basename(src)
        dst = os.path.join(dest, fn)
        shutil.move(src, dst)
        moved.append(fn)

    # Usuń pusty podfolder faktury jeśli taki był
    for src in files_to_move:
        parent = os.path.dirname(src)
        if parent != BASE_DIR:
            try:
                if not os.listdir(parent):
                    os.rmdir(parent)
            except Exception:
                pass

    prog = load_progress()
    prog[num] = {'verified': True}
    save_progress(prog)
    _cache = None  # wyczyść cache

    # Znajdź następną niezweryfikowaną (pomijaj typ 'inne')
    all_nums = sorted(inv.keys())
    idx      = all_nums.index(num) if num in all_nums else -1
    nxt      = None
    for i in list(range(idx + 1, len(all_nums))) + list(range(0, idx)):
        candidate = all_nums[i]
        if not prog.get(candidate, {}).get('verified', False) and inv[candidate]['type'] != 'inne':
            nxt = candidate
            break

    return jsonify({'ok': True, 'moved': moved, 'next': nxt})

@app.route('/api/unverify/<num>', methods=['POST'])
def api_unverify(num):
    global _cache
    inv = load_invoices()
    if num not in inv:
        abort(404)

    inv_type  = inv[num]['type']
    src_dir   = os.path.join(VERIFIED_BASE, 'Inne' if inv_type == 'inne' else inv_type)

    # Przenieś pliki z powrotem do BASE_DIR
    moved = []
    if os.path.isdir(src_dir):
        for fn in list(os.listdir(src_dir)):
            if num in fn and (fn.lower().endswith('.pdf') or fn.lower().endswith('.xml')):
                src = os.path.join(src_dir, fn)
                dst = os.path.join(BASE_DIR, fn)
                shutil.move(src, dst)
                moved.append(fn)

    prog = load_progress()
    if num in prog:
        prog[num]['verified'] = False
    save_progress(prog)
    _cache = None

    return jsonify({'ok': True, 'moved': moved})

@app.route('/api/star/<num>', methods=['POST'])
def api_star(num):
    prog = load_progress()
    if num not in prog:
        prog[num] = {}
    prog[num]['starred'] = not prog[num].get('starred', False)
    save_progress(prog)
    return jsonify({'ok': True, 'starred': prog[num]['starred']})

@app.route('/api/note/<num>', methods=['POST'])
def api_note(num):
    note = request.json.get('note', '')
    prog = load_progress()
    if num not in prog:
        prog[num] = {}
    prog[num]['note'] = note
    save_progress(prog)
    return jsonify({'ok': True})

@app.route('/api/refresh')
def api_refresh():
    global _cache
    _cache = None
    extracted = extract_zips(BASE_DIR)  # wypakuj przed przeładowaniem
    inv  = load_invoices(force=True)
    prog = load_progress()
    return jsonify({
        'count':     len(inv),
        'verified':  sum(1 for n in inv if prog.get(n, {}).get('verified', False)),
        'extracted': extracted,
    })

@app.route('/pdf/<path:filename>')
def serve_pdf(filename):
    # Szukaj rekurencyjnie w folderze głównym (obsługa podfolderów)
    verified_abs = os.path.abspath(VERIFIED_BASE)
    for dirpath, dirnames, filenames in os.walk(BASE_DIR):
        if os.path.abspath(dirpath).startswith(verified_abs):
            dirnames.clear()
            continue
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        if filename in filenames:
            return send_file(os.path.join(dirpath, filename), mimetype='application/pdf')
    # Szukaj w folderach zweryfikowanych
    for t in VERIFIED_SUBDIRS:
        p = os.path.join(VERIFIED_BASE, t, filename)
        if os.path.exists(p):
            return send_file(p, mimetype='application/pdf')
    abort(404)

@app.route('/api/check/<num>', methods=['POST'])
def api_check(num):
    """Zapisuje stan checkboxu dla WZ lub EVOUCHER."""
    data  = request.get_json()
    kind  = data.get('kind')   # 'wz' lub 'ev'
    key   = data.get('key')
    value = data.get('value')  # True/False
    prog  = load_progress()
    if num not in prog:
        prog[num] = {}
    section = 'wz_checked' if kind == 'wz' else 'ev_checked'
    if section not in prog[num]:
        prog[num][section] = {}
    prog[num][section][key] = value
    save_progress(prog)
    return jsonify({'ok': True})

# ─── HTML ─────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1280">
<title>Weryfikator faktur</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700&family=Geist+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  *,*::before,*::after { box-sizing: border-box; }
  html, body { height: 100%; margin: 0; }
  :root {
    --row-h: 44px;
    --side-w: 280px;
    --bg:        oklch(0.975 0.003 240);
    --bg-deep:   oklch(0.96 0.004 240);
    --card:      #ffffff;
    --line:      oklch(0.92 0.005 240);
    --line-2:    oklch(0.88 0.006 240);
    --ink:       oklch(0.22 0.012 250);
    --ink-2:     oklch(0.42 0.012 250);
    --ink-3:     oklch(0.6  0.012 250);
    --ink-4:     oklch(0.74 0.008 240);
    --green:     #0f8a5f;
    --green-2:   color-mix(in oklab, var(--green) 82%, white);
    --green-bg:  color-mix(in oklab, var(--green) 8%, white);
    --green-line:color-mix(in oklab, var(--green) 30%, white);
    --blue:      oklch(0.55 0.17 250);
    --blue-bg:   oklch(0.965 0.04 250);
    --amber:     oklch(0.62 0.14 60);
    --amber-bg:  oklch(0.97 0.04 70);
    --violet:    oklch(0.55 0.18 295);
    --violet-bg: oklch(0.965 0.04 295);
    --radius: 10px;
    --shadow-card: 0 1px 0 oklch(1 0 0 / 0.7) inset, 0 1px 2px oklch(0 0 0 / 0.04), 0 4px 12px oklch(0 0 0 / 0.03);
  }
  body {
    background: var(--bg);
    color: var(--ink);
    font-family: 'Geist', ui-sans-serif, system-ui, -apple-system, 'Segoe UI', sans-serif;
    font-size: 13px;
    line-height: 1.4;
    -webkit-font-smoothing: antialiased;
    overflow: hidden;
  }
  .mono { font-family: 'Geist Mono', ui-monospace, 'Segoe UI Mono', monospace; font-variant-numeric: tabular-nums; }

  /* ── App shell ── */
  .app { display: grid; grid-template-columns: var(--side-w) 1fr; height: 100vh; width: 100vw; min-width: 960px; }

  /* ── Sidebar ── */
  .side { display: flex; flex-direction: column; background: var(--card); border-right: 1px solid var(--line); min-height: 0; }
  .side-hd { padding: 16px 16px 12px; border-bottom: 1px solid var(--line); }
  .brand { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
  .brand-mark { width: 22px; height: 22px; border-radius: 6px; background: linear-gradient(135deg, var(--ink) 0%, oklch(0.35 0.02 250) 100%); display: grid; place-items: center; color: white; font-weight: 700; font-size: 11px; font-family: 'Geist Mono', monospace; }
  .brand-title { font-weight: 600; font-size: 13px; }
  .brand-sub { font-size: 11px; color: var(--ink-3); margin-top: 1px; }
  .progress-wrap { display: flex; flex-direction: column; gap: 6px; }
  .progress-row { display: flex; align-items: baseline; justify-content: space-between; font-size: 11px; }
  .progress-row .lhs { color: var(--ink-2); font-weight: 500; }
  .progress-row .rhs { color: var(--ink-3); }
  .progress-bar { height: 4px; border-radius: 2px; background: var(--bg-deep); overflow: hidden; }
  .progress-fill { height: 100%; background: var(--green); border-radius: 2px; transition: width 0.4s cubic-bezier(0.2,0.8,0.2,1); }

  /* Tabs */
  .tabs { display: flex; flex-wrap: wrap; gap: 2px; padding: 8px 10px 6px; border-bottom: 1px solid var(--line); }
  .tab { appearance: none; border: 0; background: transparent; padding: 4px 7px; border-radius: 6px; font: inherit; font-size: 11.5px; font-weight: 500; color: var(--ink-2); cursor: default; display: inline-flex; align-items: center; gap: 5px; }
  .tab:hover { background: var(--bg-deep); color: var(--ink); }
  .tab.active { background: var(--ink); color: white; }
  .tab .cnt { font-size: 10px; opacity: 0.6; font-family: 'Geist Mono', monospace; }
  .tab.active .cnt { opacity: 0.7; }

  /* Search */
  .search-wrap { padding: 8px 12px 6px; position: relative; }
  .search { width: 100%; height: 30px; padding: 0 26px 0 30px; border-radius: 7px; border: 1px solid var(--line); background: var(--bg-deep); color: var(--ink); font: inherit; font-size: 12px; outline: none; transition: border-color 0.15s, background 0.15s; box-sizing: border-box; }
  .search:focus { border-color: var(--blue); background: var(--card); box-shadow: 0 0 0 3px oklch(0.55 0.17 250 / 0.12); }
  .search::placeholder { color: var(--ink-3); }
  .search-icon { position: absolute; left: 22px; top: 50%; transform: translateY(-50%); color: var(--ink-3); pointer-events: none; }
  .search-clear { position: absolute; right: 22px; top: 50%; transform: translateY(-50%); color: var(--ink-3); background: none; border: none; cursor: pointer; font-size: 14px; line-height: 1; padding: 2px; display: none; }
  .search-clear:hover { color: var(--ink); }
  .search-wrap:has(input:not(:placeholder-shown)) .search-clear { display: block; }

  /* Filter bar */
  .filter-bar { padding: 6px 12px 8px; border-bottom: 1px solid var(--line); display: flex; flex-direction: column; gap: 5px; }
  .filter-row { display: flex; align-items: center; gap: 5px; }
  .filter-lbl { font-size: 11px; color: var(--ink-3); white-space: nowrap; }
  .filter-date { height: 24px; padding: 0 5px; border-radius: 5px; border: 1px solid var(--line); background: var(--bg-deep); color: var(--ink); font: inherit; font-size: 11px; outline: none; flex: 1; min-width: 0; }
  .filter-date:focus { border-color: var(--blue); }
  .filter-select { height: 24px; padding: 0 4px; border-radius: 5px; border: 1px solid var(--line); background: var(--bg-deep); color: var(--ink); font: inherit; font-size: 11px; outline: none; flex: 1; }
  .filter-select:focus { border-color: var(--blue); }
  .filter-clear { appearance: none; border: 0; background: transparent; color: var(--ink-3); font-size: 12px; cursor: default; padding: 2px 4px; border-radius: 4px; }
  .filter-clear:hover { background: var(--bg-deep); color: var(--ink); }
  .filter-star-lbl { cursor: default; display: flex; align-items: center; gap: 3px; margin-left: 4px; white-space: nowrap; }

  /* Type sums */
  .type-sums { border-bottom: 1px solid var(--line); }
  .type-sum-row { display: flex; justify-content: space-between; align-items: baseline; padding: 5px 14px; font-size: 11.5px; }
  .type-sum-row + .type-sum-row { border-top: 1px solid var(--line); }
  .type-sum-lbl { color: var(--ink-3); }
  .type-sum-val { font-family: 'Geist Mono', monospace; font-weight: 600; color: var(--ink); font-size: 12px; }

  /* List */
  .list { flex: 1; min-height: 0; overflow-y: auto; padding: 4px 0; scrollbar-width: thin; scrollbar-color: var(--line-2) transparent; }
  .list::-webkit-scrollbar { width: 6px; }
  .list::-webkit-scrollbar-thumb { background: var(--line-2); border-radius: 3px; }

  .item { display: grid; grid-template-columns: 10px auto 1fr auto auto; align-items: center; column-gap: 10px; padding: 9px 14px 9px 12px; border-left: 2px solid transparent; cursor: default; transition: background 0.1s; }
  .item:hover { background: var(--bg-deep); }
  .item.selected { background: var(--blue-bg); border-left-color: var(--blue); }
  .item.verified { opacity: 0.45; }
  .item.verified:hover { opacity: 0.75; }
  .item-dot { width: 6px; height: 6px; border-radius: 50%; background: transparent; border: 1.5px solid var(--ink-4); }
  .item.verified .item-dot { background: var(--green); border-color: var(--green); }

  .item-badge { display: inline-flex; align-items: center; justify-content: center; height: 18px; min-width: 22px; padding: 0 5px; border-radius: 4px; font-family: 'Geist Mono', monospace; font-size: 10px; font-weight: 600; }
  .badge-63       { background: var(--green-bg);  color: var(--green); }
  .badge-66       { background: var(--amber-bg);  color: var(--amber); }
  .badge-17       { background: var(--violet-bg); color: var(--violet); }
  .badge-15       { background: oklch(0.92 0.05 200); color: oklch(0.35 0.12 200); }
  .badge-billbird { background: oklch(0.92 0.06 30);  color: oklch(0.40 0.14 30); }
  .badge-inne     { background: oklch(0.94 0.003 240); color: var(--ink-3); }

  .item-main { min-width: 0; display: flex; flex-direction: column; gap: 1px; }
  .item-num { font-family: 'Geist Mono', monospace; font-size: 12px; font-weight: 500; color: var(--ink); }
  .item-meta { font-size: 11px; color: var(--ink-3); font-family: 'Geist Mono', monospace; }
  .item-meta .dot { padding: 0 3px; opacity: 0.5; }
  .item-check { color: var(--green); opacity: 0; transition: opacity 0.15s; }
  .item.verified .item-check { opacity: 1; }
  .item-star { font-size: 13px; line-height: 1; opacity: 0; transition: opacity 0.15s; }
  .item.starred .item-star { opacity: 1; }
  .item:hover .item-star { opacity: 0.35; }
  .item.starred:hover .item-star { opacity: 1; }

  /* Note textarea */
  .note-wrap { padding: 12px 20px 0; }
  .note-area { width: 100%; min-height: 60px; max-height: 140px; padding: 8px 10px; border-radius: 7px; border: 1px solid var(--line); background: var(--bg-deep); color: var(--ink); font: inherit; font-size: 12.5px; resize: vertical; outline: none; transition: border-color 0.15s; box-sizing: border-box; }
  .note-area:focus { border-color: var(--blue); background: var(--card); box-shadow: 0 0 0 3px oklch(0.55 0.17 250 / 0.12); }
  .note-area::placeholder { color: var(--ink-3); }

  /* Sidebar footer */
  .side-ft { padding: 8px 14px; border-top: 1px solid var(--line); font-size: 10.5px; color: var(--ink-3); display: flex; flex-wrap: wrap; gap: 8px 12px; align-items: center; background: var(--bg-deep); }
  .kbd { display: inline-flex; align-items: center; padding: 1px 5px; border-radius: 4px; background: white; border: 1px solid var(--line); font-family: 'Geist Mono', monospace; font-size: 10px; color: var(--ink-2); box-shadow: 0 1px 0 var(--line); }
  .side-ft .grp { display: inline-flex; gap: 4px; align-items: center; }

  /* ── Main ── */
  .main { overflow-y: auto; padding: 24px 28px 60px; scrollbar-width: thin; scrollbar-color: var(--line-2) transparent; }
  .main::-webkit-scrollbar { width: 8px; }
  .main::-webkit-scrollbar-thumb { background: var(--line-2); border-radius: 4px; border: 2px solid transparent; background-clip: content-box; }
  .empty-state { display: flex; align-items: center; justify-content: center; height: 60vh; color: var(--ink-3); font-size: 14px; }

  /* ── Card ── */
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; box-shadow: var(--shadow-card); max-width: 1060px; margin: 0 auto; }

  /* Detail header */
  .det-hd { display: flex; align-items: center; justify-content: space-between; padding: 18px 24px; border-bottom: 1px solid var(--line); }
  .det-hd-l { display: flex; align-items: center; gap: 12px; }
  .det-hd-badge { height: 24px; min-width: 30px; padding: 0 8px; border-radius: 6px; display: inline-flex; align-items: center; justify-content: center; font-family: 'Geist Mono', monospace; font-size: 12px; font-weight: 600; }
  .det-hd-num { font-family: 'Geist Mono', monospace; font-size: 22px; font-weight: 600; letter-spacing: -0.01em; color: var(--ink); border-radius: 4px; padding: 0 4px; transition: background 0.15s; }
  .det-hd-num:hover { background: var(--surface-2); }
  .det-hd-sub { font-size: 11.5px; color: var(--ink-3); }
  .det-hd-r { display: flex; align-items: center; gap: 8px; }
  .det-hd-chip { display: inline-flex; align-items: center; gap: 5px; padding: 3px 9px; border-radius: 999px; background: var(--bg-deep); color: var(--ink-2); font-size: 11px; }

  /* Metadata grid */
  .meta { display: grid; gap: 0; border-bottom: 1px solid var(--line); }
  .meta-cell { padding: 14px 20px 16px; border-right: 1px solid var(--line); display: flex; flex-direction: column; gap: 6px; }
  .meta-cell:last-child { border-right: 0; }
  .meta-lbl { font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-3); }
  .meta-val { font-family: 'Geist Mono', monospace; font-size: 15px; font-weight: 500; color: var(--ink); }
  .meta-val.money { color: var(--green); font-weight: 600; }
  .meta-val.money-due { color: var(--amber); font-weight: 600; }

  /* Section header */
  .sec-hd { display: flex; align-items: center; justify-content: space-between; padding: 14px 24px 8px; }
  .sec-hd-l { display: flex; align-items: baseline; gap: 8px; }
  .sec-hd-title { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-3); }
  .sec-hd-count { font-family: 'Geist Mono', monospace; font-size: 11px; color: var(--ink-3); }
  .sec-hd-progress { font-size: 11px; color: var(--green); font-family: 'Geist Mono', monospace; font-weight: 500; }

  /* WZ list */
  .wz-list { padding: 0 12px 8px; display: flex; flex-direction: column; }
  .wz-row { display: grid; grid-template-columns: 28px 1fr; align-items: center; column-gap: 10px; padding: 7px 12px; border-radius: 8px; transition: background 0.1s; }
  .wz-row:hover { background: var(--bg-deep); }
  .wz-row.checked { background: var(--green-bg); }
  .wz-row.checked:hover { background: color-mix(in oklab, var(--green) 12%, white); }
  .wz-row.focused { outline: 2px solid var(--blue); outline-offset: -2px; }
  .wz-row.focused.checked { outline-color: var(--green); }

  /* Checkbox button */
  .check { width: 18px; height: 18px; border-radius: 5px; border: 1.5px solid var(--line-2); background: white; display: grid; place-items: center; cursor: default; transition: all 0.12s; flex-shrink: 0; appearance: none; }
  .check:hover { border-color: var(--ink-3); }
  .check.on { background: var(--green); border-color: var(--green); }
  .check.on::after { content: ''; width: 10px; height: 6px; border-left: 2px solid white; border-bottom: 2px solid white; transform: rotate(-45deg) translate(1px, -1px); display: block; }

  .wz-num { font-family: 'Geist Mono', monospace; font-size: 13.5px; font-weight: 500; color: var(--blue); cursor: default; display: inline-flex; align-items: center; gap: 8px; }
  .wz-num:hover { text-decoration: underline; text-underline-offset: 3px; }
  .wz-row.checked .wz-num { color: var(--green); }
  .wz-copy-hint { font-size: 10px; color: var(--ink-3); opacity: 0; transition: opacity 0.15s; font-family: 'Geist Mono', monospace; }
  .wz-row:hover .wz-copy-hint { opacity: 0.8; }

  /* EVOUCHER table */
  .ev-table { width: calc(100% - 24px); margin: 0 12px 12px; border-collapse: collapse; }
  .ev-table th { text-align: left; padding: 8px 12px; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-3); border-bottom: 1px solid var(--line); }
  .ev-table th.num { text-align: right; }
  .ev-table td { padding: 8px 12px; border-bottom: 1px solid var(--line); font-size: 13px; }
  .ev-table td.num { text-align: right; font-family: 'Geist Mono', monospace; }
  .ev-table tr:last-child td { border-bottom: 0; }
  .ev-table tr.checked { background: var(--green-bg); }
  .ev-table tr.checked td.total { color: var(--green); font-weight: 600; }
  .ev-name { font-weight: 500; color: var(--ink); cursor: default; }
  .ev-name:hover { text-decoration: underline; color: var(--blue); }

  /* Actions */
  .actions { display: flex; gap: 8px; padding: 16px 20px; border-top: 1px solid var(--line); background: oklch(0.985 0.003 240); border-radius: 0 0 12px 12px; }
  .btn { appearance: none; border: 0; height: 38px; padding: 0 16px; border-radius: 8px; font: inherit; font-size: 13px; font-weight: 500; cursor: default; display: inline-flex; align-items: center; justify-content: center; gap: 7px; transition: all 0.12s; }
  .btn-ghost { background: white; border: 1px solid var(--line); color: var(--ink); }
  .btn-ghost:hover { background: var(--bg-deep); border-color: var(--line-2); }
  .btn-ghost.active { background: var(--ink); color: white; border-color: var(--ink); }
  .btn-primary { flex: 1; background: var(--green); color: white; font-weight: 600; box-shadow: 0 1px 0 rgba(255,255,255,0.2) inset; }
  .btn-primary:hover { background: var(--green-2); }
  .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-danger { background: var(--amber-bg); color: var(--amber); border: 1px solid color-mix(in oklab, var(--amber) 25%, transparent); font-weight: 600; }
  .btn-danger:hover { background: color-mix(in oklab, var(--amber) 14%, white); }
  .kbd-inline { margin-left: 4px; padding: 1px 5px; border-radius: 3px; background: oklch(1 0 0 / 0.18); font-family: 'Geist Mono', monospace; font-size: 10px; font-weight: 600; }
  .btn-ghost .kbd-inline { background: var(--bg-deep); color: var(--ink-3); border: 1px solid var(--line); }
  .btn-danger .kbd-inline { background: color-mix(in oklab, var(--amber) 15%, transparent); color: var(--amber); }

  /* PDF */
  .pdf-wrap { display: none; border-top: 1px solid var(--line); height: 52vh; border-radius: 0 0 12px 12px; overflow: hidden; }
  .pdf-wrap.pdf-open { display: block; }
  .pdf-wrap iframe { width: 100%; height: 100%; border: none; display: block; }

  /* Icons */
  .ico { width: 14px; height: 14px; display: inline-block; vertical-align: -2px; flex-shrink: 0; }

  /* Toast */
  .toast { position: fixed; right: 24px; bottom: 24px; background: var(--ink); color: white; padding: 10px 14px; border-radius: 10px; font-size: 12.5px; font-weight: 500; display: flex; align-items: center; gap: 10px; box-shadow: 0 10px 30px oklch(0 0 0 / 0.25); z-index: 1000; animation: toast-in 0.22s cubic-bezier(0.2,0.8,0.2,1); }
  .toast-icon { width: 18px; height: 18px; border-radius: 50%; background: var(--green-2); display: grid; place-items: center; flex-shrink: 0; }
  .toast-icon::after { content: ''; width: 8px; height: 5px; border-left: 2px solid white; border-bottom: 2px solid white; transform: rotate(-45deg) translate(1px,-1px); display: block; }
  @keyframes toast-in { from { opacity:0; transform:translateY(8px) scale(0.98); } to { opacity:1; transform:translateY(0) scale(1); } }
</style>
</head>
<body>

<div class="app">
  <!-- SIDEBAR -->
  <aside class="side">
    <div class="side-hd">
      <div class="brand">
        <div class="brand-mark">FV</div>
        <div>
          <div class="brand-title">Weryfikator faktur</div>
          <div class="brand-sub" id="brand-sub">Ładowanie…</div>
        </div>
      </div>
      <div class="progress-wrap">
        <div class="progress-row">
          <span class="lhs" id="prog-text">—</span>
          <span class="rhs mono" id="prog-pct">0%</span>
        </div>
        <div class="progress-bar">
          <div class="progress-fill" id="prog-fill" style="width:0%"></div>
        </div>
      </div>
    </div>
    <div class="tabs">
      <button class="tab active" onclick="setFilter('all')">Wszystkie<span class="cnt" id="cnt-all">0</span></button>
      <button class="tab" onclick="setFilter('63')">63<span class="cnt" id="cnt-63">0</span></button>
      <button class="tab" onclick="setFilter('66')">66<span class="cnt" id="cnt-66">0</span></button>
      <button class="tab" onclick="setFilter('17')">17<span class="cnt" id="cnt-17">0</span></button>
      <button class="tab" onclick="setFilter('15')">15<span class="cnt" id="cnt-15">0</span></button>
      <button class="tab" onclick="setFilter('billbird')">BB<span class="cnt" id="cnt-billbird">0</span></button>
      <button class="tab" onclick="setFilter('inne')">Inne<span class="cnt" id="cnt-inne">0</span></button>
    </div>
    <div class="search-wrap">
      <span class="search-icon">
        <svg class="ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round">
          <circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5L14 14"/>
        </svg>
      </span>
      <input class="search" type="text" id="q" placeholder="Szukaj numeru, WZ, kwoty lub daty…" oninput="renderList()">
      <button class="search-clear" onclick="clearSearch()" tabindex="-1">✕</button>
    </div>
    <div class="filter-bar">
      <div class="filter-row">
        <label class="filter-lbl">Od</label><input class="filter-date" type="date" id="date-from" onchange="renderList()">
        <label class="filter-lbl">Do</label><input class="filter-date" type="date" id="date-to" onchange="renderList()">
        <button class="filter-clear" onclick="clearDates()" title="Wyczyść daty">✕</button>
      </div>
      <div class="filter-row">
        <label class="filter-lbl">Sortuj</label>
        <select class="filter-select" id="sort-by" onchange="renderList()">
          <option value="num">Numer</option>
          <option value="date">Data</option>
          <option value="amount">Kwota</option>
          <option value="type">Typ</option>
        </select>
        <label class="filter-lbl filter-star-lbl">
          <input type="checkbox" id="only-starred" onchange="renderList()">⭐ Gwiazdki
        </label>
      </div>
    </div>
    <div class="list" id="inv-list"></div>
    <div class="side-ft">
      <span class="grp"><span class="kbd">J</span><span class="kbd">K</span> faktura</span>
      <span class="grp"><span class="kbd">N</span><span class="kbd">M</span> WZ</span>
      <span class="grp"><span class="kbd">Spacja</span> zaznacz</span>
      <span class="grp"><span class="kbd">V</span> zatwierdź</span>
      <span class="grp"><span class="kbd">U</span> cofnij</span>
      <span class="grp"><span class="kbd">P</span> PDF</span>
    </div>
  </aside>

  <!-- MAIN -->
  <main class="main" id="main">
    <div class="empty-state">← Wybierz fakturę z listy</div>
  </main>
</div>

<script>
let all          = [];
let filter       = 'all';
let current      = null;
let verifyStack  = [];     // stos zatwierdzonych — U cofa ostatnią z góry
let pdfOpen      = false;
let pdfFile      = null;
let toastTimer   = null;

// ── SVG icons ─────────────────────────────────
const ICO = {
  check:   `<svg class="ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8.5L6.5 12L13 4.5"/></svg>`,
  refresh: `<svg class="ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3v4h-4M2 13v-4h4"/><path d="M3.5 7a5 5 0 0 1 8.6-1.5L14 7M12.5 9a5 5 0 0 1-8.6 1.5L2 9"/></svg>`,
  file:    `<svg class="ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 1.5H4a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5.5L9 1.5z"/><path d="M9 1.5V5.5h4"/></svg>`,
  undo:    `<svg class="ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6H9.5a3.5 3.5 0 0 1 0 7H7M3 6l3-3M3 6l3 3"/></svg>`,
};

// ── Init ──────────────────────────────────────
async function init() {
  const r = await fetch('/api/invoices');
  all = await r.json();
  renderList();
  updateProgress();
}

function fmtMoney(v) {
  const s = v.toFixed(2), [int, frac] = s.split('.');
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + frac + ' PLN';
}

function updateProgress() {
  const main  = all.filter(i => i.type !== 'inne');
  const total = main.length;
  const done  = main.filter(i => i.verified).length;
  const pct   = total ? Math.round(done / total * 100) : 0;
  document.getElementById('prog-text').textContent = `${done}/${total} zweryfikowanych`;
  document.getElementById('prog-pct').textContent  = pct + '%';
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('brand-sub').textContent = `${total} dokumentów`;
  document.getElementById('cnt-all').textContent  = total;
  document.getElementById('cnt-63').textContent   = all.filter(i=>i.type==='63').length;
  document.getElementById('cnt-66').textContent   = all.filter(i=>i.type==='66').length;
  document.getElementById('cnt-17').textContent       = all.filter(i=>i.type==='17').length;
  document.getElementById('cnt-15').textContent       = all.filter(i=>i.type==='15').length;
  document.getElementById('cnt-billbird').textContent = all.filter(i=>i.type==='billbird').length;
  document.getElementById('cnt-inne').textContent     = all.filter(i=>i.type==='inne').length;
}

// ── Wyszukiwanie kwot ─────────────────────────
function matchesAmounts(q, amounts) {
  if (!amounts || !amounts.length) return false;
  // Normalizuj zapytanie: usuń spacje, zamień przecinek na kropkę
  const qn = q.replace(/\s/g, '').replace(',', '.');
  const qf = parseFloat(qn);
  if (isNaN(qf) || qf <= 0) return false;
  for (const a of amounts) {
    if (!a) continue;
    // Dopasowanie dokładne (float z tolerancją 0.005 zł)
    const af = parseFloat(a);
    if (!isNaN(af) && Math.abs(qf - af) < 0.005) return true;
    // Dopasowanie prefiksowe: "1366" pasuje do "1366.09"
    if (a.startsWith(qn)) return true;
  }
  return false;
}

// ── Filtry i lista ────────────────────────────
function setFilter(f) {
  filter = f;
  document.querySelectorAll('.tab').forEach((t, i) => {
    t.classList.toggle('active', ['all','63','66','17','15','billbird','inne'][i] === f);
  });
  renderList();
}

function clearSearch() {
  const q = document.getElementById('q');
  q.value = '';
  q.focus();
  renderList();
}

function clearDates() {
  document.getElementById('date-from').value = '';
  document.getElementById('date-to').value   = '';
  renderList();
}

function getVisible() {
  const q        = document.getElementById('q').value.trim().toLowerCase();
  const dateFrom = document.getElementById('date-from').value;   // YYYY-MM-DD
  const dateTo   = document.getElementById('date-to').value;
  const sortBy   = document.getElementById('sort-by').value;
  const onlyStar = document.getElementById('only-starred').checked;

  let items = all;

  // Filtr zakładki
  if (filter !== 'all') {
    items = items.filter(i => i.type === filter);
  } else {
    if (!q && !dateFrom && !dateTo) items = items.filter(i => i.type !== 'inne');
  }

  // Filtr gwiazdek
  if (onlyStar) items = items.filter(i => i.starred);

  // Filtr dat od-do
  if (dateFrom) items = items.filter(i => i.issue_date_iso >= dateFrom);
  if (dateTo)   items = items.filter(i => i.issue_date_iso <= dateTo);

  // Wyszukiwanie tekstowe
  if (q) items = items.filter(i =>
    i.invoice_num.includes(q) ||
    (i.wz_numbers && i.wz_numbers.some(wz => wz.includes(q))) ||
    matchesAmounts(q, i.amounts_raw) ||
    matchesDate(q, i.issue_date_iso) ||
    (i.numer_faktury_korygowanej && i.numer_faktury_korygowanej.toLowerCase().includes(q))
  );

  // Sortowanie
  items = [...items];
  if (sortBy === 'date') {
    items.sort((a, b) => (b.issue_date_iso || '').localeCompare(a.issue_date_iso || ''));
  } else if (sortBy === 'amount') {
    items.sort((a, b) => parseFloat(b.total_raw || 0) - parseFloat(a.total_raw || 0));
  } else if (sortBy === 'type') {
    items.sort((a, b) => a.type.localeCompare(b.type) || a.invoice_num.localeCompare(b.invoice_num));
  } else {
    items.sort((a, b) => a.invoice_num.localeCompare(b.invoice_num));
  }

  // Niezweryfikowane na górze
  return [...items.filter(i => !i.verified), ...items.filter(i => i.verified)];
}

function matchesDate(q, iso) {
  if (!iso || q.length < 3) return false;
  // Akceptuj formaty: '2026', '03.2026', '03/2026', '2026-03', '25.03', '25.03.2026'
  const norm = q.replace(/\./g, '-').replace(/\//g, '-');
  // Rok: '2026' → pasuje do '2026-xx-xx'
  if (/^\d{4}$/.test(q)) return iso.startsWith(q);
  // Rok-miesiąc: '2026-03'
  if (/^\d{4}-\d{2}$/.test(norm)) return iso.startsWith(norm);
  // Miesiąc-rok: '03-2026' → zmień na '2026-03'
  const mm = norm.match(/^(\d{2})-(\d{4})$/);
  if (mm) return iso.startsWith(`${mm[2]}-${mm[1]}`);
  // Dzień-miesiąc: '25-03' → sprawdź czy iso zawiera '-03-25'
  const dm = norm.match(/^(\d{2})-(\d{2})$/);
  if (dm) return iso.includes(`-${dm[2]}-${dm[1]}`);
  return false;
}

function renderList() {
  const list  = document.getElementById('inv-list');
  const items = getVisible();
  if (!items.length) {
    list.innerHTML = `<div style="padding:24px 16px;text-align:center;color:var(--ink-3);font-size:12px">Brak wyników</div>`;
    return;
  }
  list.innerHTML = items.map(i => `
    <div class="item ${i.verified?'verified':''} ${i.starred?'starred':''} ${i.invoice_num===current?'selected':''}"
         onclick="load('${i.invoice_num}')">
      <span class="item-dot"></span>
      <span class="item-badge badge-${i.type}">${i.type==='billbird'?'BB':i.type==='inne'?'?':i.type}</span>${i.has_korekta?'<span class="item-badge" style="background:oklch(0.92 0.06 0);color:oklch(0.40 0.14 0);margin-left:2px" title="Istnieje faktura korygująca">K</span>':''}
      <span class="item-main">
        <span class="item-num">${i.invoice_num}</span>
        <span class="item-meta">${i.issue_date}<span class="dot"> · </span>${i.total||'—'}</span>
      </span>
      <span class="item-star" onclick="event.stopPropagation();toggleStar('${i.invoice_num}')">⭐</span>
      <span class="item-check">${ICO.check}</span>
    </div>`).join('');
  // scroll selected into view
  const sel = list.querySelector('.selected');
  if (sel) {
    const lb = list.getBoundingClientRect(), sb = sel.getBoundingClientRect();
    if (sb.top < lb.top + 36 || sb.bottom > lb.bottom - 36)
      list.scrollTop += (sb.top - lb.top) - lb.height/2 + sb.height/2;
  }
}

// ── Navigate list ─────────────────────────────
function navigateList(dir) {
  const visible = getVisible();
  if (!visible.length) return;
  const idx = visible.findIndex(i => i.invoice_num === current);
  let next = Math.max(0, Math.min(visible.length - 1, idx + dir));
  if (next !== idx) load(visible[next].invoice_num);
}

// ── Gwiazdka ──────────────────────────────────
async function toggleStar(num) {
  const r    = await fetch(`/api/star/${num}`, {method:'POST'});
  const data = await r.json();
  const inv  = all.find(i => i.invoice_num === num);
  if (inv) inv.starred = data.starred;
  renderList();
  // Odśwież przycisk gwiazdki w nagłówku jeśli to aktualna faktura
  const btn = document.getElementById('btn-star');
  if (btn && num === current) btn.textContent = data.starred ? '⭐ Gwiazdka' : '☆ Gwiazdka';
}

// ── Notatka ───────────────────────────────────
let noteTimer = null;
function scheduleNoteSave(num) {
  clearTimeout(noteTimer);
  noteTimer = setTimeout(() => saveNote(num), 800);
}
async function saveNote(num) {
  const el = document.getElementById('note-area');
  if (!el) return;
  await fetch(`/api/note/${num}`, {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({note: el.value})
  });
}

// ── WZ nawigacja ──────────────────────────────
let wzFocusIdx = -1;

function wzNav(dir) {
  const rows = [...document.querySelectorAll('.wz-row')];
  if (!rows.length) return;
  wzFocusIdx = Math.max(0, Math.min(rows.length - 1, wzFocusIdx + dir));
  rows.forEach((r, i) => r.classList.toggle('focused', i === wzFocusIdx));
  rows[wzFocusIdx].scrollIntoView({block: 'nearest'});
}

function wzCheck(num) {
  const rows = [...document.querySelectorAll('.wz-row')];
  if (!rows.length) return;
  // Jeśli żaden nie jest podświetlony — zacznij od pierwszego
  if (wzFocusIdx < 0 || wzFocusIdx >= rows.length) wzFocusIdx = 0;
  const row = rows[wzFocusIdx];
  const btn = row.querySelector('.check');
  const wz  = row.querySelector('.wz-num').textContent.replace('kopiuj','').trim();
  btn.click();
}

// ── Przejdź do faktury korygującej ───────────
function showKorekta(num) {
  const q = document.getElementById('q');
  q.value = num;
  setFilter('all');
  renderList();
  q.focus();
}

// ── Ładowanie faktury ─────────────────────────
async function load(num) {
  current = num;
  renderList();
  const r = await fetch(`/api/invoice/${num}`);
  const d = await r.json();
  pdfFile    = d.pdf_filename || null;
  wzFocusIdx = -1;

  const wzChecked = d.wz_checked || {};
  const evChecked = d.ev_checked || {};
  const typeNames = {'63':'Dostawa towaru','66':'Wyroby tytoniowe','17':'Vouchery cyfrowe','15':'Farmacja','billbird':'BillBird','inne':'Inny typ'};

  // ── Header ──
  let html = `<div class="card"><div class="det-hd">
    <div class="det-hd-l">
      <span class="det-hd-badge badge-${d.type}">${d.type}</span>
      <span class="det-hd-num" onclick="copyToClipboard('${d.invoice_num}')" title="Kliknij aby skopiować numer faktury" style="cursor:pointer">${d.invoice_num}</span>
      <span class="det-hd-sub">· ${typeNames[d.type]||''}</span>
    </div>
    <div class="det-hd-r">`;
  html += `<button class="btn btn-ghost" id="btn-star" onclick="toggleStar('${num}')" style="font-size:13px;padding:4px 8px">${d.starred ? '⭐ Gwiazdka' : '☆ Gwiazdka'}</button>`;
  const listItem = all.find(i => i.invoice_num === num);
  if (listItem && listItem.has_korekta)
    html += `<button class="btn btn-ghost" onclick="showKorekta('${num}')" style="color:oklch(0.40 0.14 0);border-color:oklch(0.75 0.1 0)" title="Pokaż fakturę korygującą">K korekta</button>`;
  if (d.verified)  html += `<span class="det-hd-chip" style="background:var(--green-bg);color:var(--green)">${ICO.check} Zweryfikowano</span>`;
  if (d.pdf_only)  html += `<span class="det-hd-chip" style="background:var(--amber-bg);color:var(--amber)">⚠ Brak XML – tylko PDF</span>`;
  html += `</div></div>`;

  // ── Meta grid ──
  if (d.pdf_only) {
    html += `<div style="padding:14px 20px;color:var(--amber);font-size:12.5px;border-bottom:1px solid var(--line);background:var(--amber-bg)">
      Plik XML jest uszkodzony lub brakuje go. Dane faktury niedostępne — sprawdź PDF.
    </div>`;
  } else {
    const hasDz  = (d.type === '63' || d.type === '15' || d.type === 'billbird') && d.do_zaplaty;
    const cols   = hasDz ? 4 : 3;
    html += `<div class="meta" style="grid-template-columns:repeat(${cols},1fr)">
      <div class="meta-cell"><div class="meta-lbl">Data wystawienia</div><div class="meta-val">${d.issue_date||'—'}</div></div>
      <div class="meta-cell"><div class="meta-lbl">Data usługi</div><div class="meta-val">${d.service_date||'—'}</div></div>
      <div class="meta-cell"><div class="meta-lbl">Kwota należności</div><div class="meta-val money">${d.total||'—'}${d.suma_obciazen ? `<span style="display:block;font-size:11px;font-weight:400;color:var(--ink-3);margin-top:2px">${d.suma_obciazen} obciążenia</span>` : ''}</div></div>
      ${hasDz ? `<div class="meta-cell"><div class="meta-lbl">Do zapłaty</div><div class="meta-val money-due">${d.do_zaplaty}</div></div>` : ''}
    </div>`;
  }

  // ── WZ section ──
  if ((d.type === '63' || d.type === '66' || d.type === '15') && d.wz_numbers && d.wz_numbers.length) {
    const done = d.wz_numbers.filter(wz => wzChecked[wz]).length;
    html += `<div class="sec-hd">
      <div class="sec-hd-l"><span class="sec-hd-title">Numery WZ</span><span class="sec-hd-count">· ${d.wz_numbers.length}</span></div>
      <div style="display:flex;align-items:center;gap:6px">
        <span class="sec-hd-progress" id="wz-prog">${done===d.wz_numbers.length ? '✓ wszystkie sprawdzone' : done+'/'+d.wz_numbers.length+' sprawdzonych'}</span>
        <button class="btn btn-ghost" style="padding:2px 8px;font-size:13px" onclick="wzNav(-1)" title="Poprzedni WZ">◀</button>
        <button class="btn btn-ghost" style="padding:2px 8px;font-size:13px" onclick="wzNav(+1)" title="Następny WZ">▶</button>
        <button class="btn btn-ghost" style="padding:2px 8px;font-size:11px" onclick="wzCheck('${num}')" title="Zaznacz / odznacz aktualny WZ">Zaznacz</button>
      </div>
    </div><div class="wz-list">`;
    d.wz_numbers.forEach(wz => {
      const checked = wzChecked[wz] || false;
      html += `<div class="wz-row${checked?' checked':''}">
        <button class="check${checked?' on':''}" onclick="toggleWz('${num}','${wz}',this)"></button>
        <span class="wz-num" onclick="copyToClipboard('${wz}')">${wz}<span class="wz-copy-hint">kopiuj</span></span>
      </div>`;
    });
    html += `</div>`;
  }

  // ── Faktura korygująca ──
  if (d.numer_faktury_korygowanej) {
    html += `<div style="padding:10px 20px;font-size:12px;color:var(--ink-2);border-bottom:1px solid var(--line)">
      Faktura korygująca do: <span class="wz-num" style="cursor:pointer" onclick="copyToClipboard('${d.numer_faktury_korygowanej}')">${d.numer_faktury_korygowanej}<span class="wz-copy-hint">kopiuj</span></span>
    </div>`;
  }

  // ── EVOUCHER section ──
  if (d.type === '17' && d.evoucher_items && d.evoucher_items.length) {
    const done = d.evoucher_items.filter(it => evChecked[it.name]).length;
    html += `<div class="sec-hd">
      <div class="sec-hd-l"><span class="sec-hd-title">Pozycje EVOUCHER</span><span class="sec-hd-count">· ${d.evoucher_items.length}</span></div>
      <span class="sec-hd-progress" id="ev-prog">${done===d.evoucher_items.length ? '✓ wszystkie sprawdzone' : done+'/'+d.evoucher_items.length+' sprawdzonych'}</span>
    </div>
    <table class="ev-table">
      <thead><tr><th style="width:40px"></th><th>Nazwa towaru</th><th class="num" style="width:80px">Ilość</th><th class="num" style="width:120px">Cena jedn.</th><th class="num" style="width:130px">Wartość</th></tr></thead>
      <tbody>`;
    d.evoucher_items.forEach(it => {
      const checked = evChecked[it.name] || false;
      const safeN   = it.name.replace(/\\/g,'\\\\').replace(/'/g,"\\'");
      html += `<tr class="${checked?'checked':''}">
        <td><button class="check${checked?' on':''}" onclick="toggleEv('${num}','${safeN}',this)"></button></td>
        <td><span class="ev-name" onclick="copyToClipboard('${safeN}')">${it.name}</span></td>
        <td class="num">${it.qty} szt.</td>
        <td class="num">${it.unit_price}</td>
        <td class="num total">${it.value}</td>
      </tr>`;
    });
    html += `</tbody></table>`;
  }

  // ── Notatka ──
  html += `<div class="note-wrap">
    <textarea id="note-area" class="note-area" placeholder="Notatka do faktury…" oninput="scheduleNoteSave('${num}')">${d.note||''}</textarea>
  </div>`;

  // ── Actions ──
  if (d.verified) {
    html += `<div class="actions">
      <button class="btn btn-ghost" onclick="refreshInvoices()">${ICO.refresh} Odśwież</button>
      <button class="btn btn-ghost${pdfOpen?' active':''}" id="btn-pdf" onclick="togglePdf()">${ICO.file} PDF<span class="kbd-inline">P</span></button>
      <button class="btn btn-danger" onclick="unverify()">${ICO.undo} Cofnij weryfikację<span class="kbd-inline">U</span></button>
    </div>`;
  } else {
    html += `<div class="actions">
      <button class="btn btn-ghost" onclick="refreshInvoices()">${ICO.refresh} Odśwież</button>
      <button class="btn btn-ghost${pdfOpen?' active':''}" id="btn-pdf" onclick="togglePdf()">${ICO.file} PDF<span class="kbd-inline">P</span></button>
      <button class="btn btn-primary" id="btn-v" onclick="verify()">${ICO.check} Zweryfikuj<span class="kbd-inline">V</span></button>
    </div>`;
  }

  // ── PDF placeholder ──
  html += `<div id="pdf-wrap" class="pdf-wrap${pdfOpen?' pdf-open':''}">
    <iframe id="pdf-frame" src=""></iframe>
  </div></div>`; // .card

  document.getElementById('main').innerHTML = html;
  // PDF-only: zawsze otwórz podgląd automatycznie
  if (d.pdf_only && pdfFile) {
    pdfOpen = true;
    document.getElementById('pdf-wrap').classList.add('pdf-open');
    const btn = document.getElementById('btn-pdf');
    if (btn) btn.classList.add('active');
    loadPdf();
  } else if (pdfOpen && pdfFile) {
    loadPdf();
  }
}

// ── Checkbox WZ ───────────────────────────────
async function toggleWz(num, wz, el) {
  const checked = !el.classList.contains('on');
  el.classList.toggle('on', checked);
  el.closest('.wz-row').classList.toggle('checked', checked);
  updateCheckedCount();
  fetch(`/api/check/${num}`, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({kind:'wz', key:wz, value:checked}) });
}

// ── Checkbox EVOUCHER ─────────────────────────
async function toggleEv(num, name, el) {
  const checked = !el.classList.contains('on');
  el.classList.toggle('on', checked);
  el.closest('tr').classList.toggle('checked', checked);
  updateCheckedCount();
  fetch(`/api/check/${num}`, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({kind:'ev', key:name, value:checked}) });
}

function updateCheckedCount() {
  const wzProg = document.getElementById('wz-prog');
  if (wzProg) {
    const total = document.querySelectorAll('.wz-row').length;
    const done  = document.querySelectorAll('.wz-row.checked').length;
    wzProg.textContent = done === total ? '✓ wszystkie sprawdzone' : `${done}/${total} sprawdzonych`;
  }
  const evProg = document.getElementById('ev-prog');
  if (evProg) {
    const total = document.querySelectorAll('.ev-table tbody tr').length;
    const done  = document.querySelectorAll('.ev-table tbody tr.checked').length;
    evProg.textContent = done === total ? '✓ wszystkie sprawdzone' : `${done}/${total} sprawdzonych`;
  }
}

// ── PDF ───────────────────────────────────────
function loadPdf() {
  const f = document.getElementById('pdf-frame');
  if (f && pdfFile) f.src = `/pdf/${encodeURIComponent(pdfFile)}`;
}
function togglePdf() {
  pdfOpen = !pdfOpen;
  const wrap = document.getElementById('pdf-wrap');
  const btn  = document.getElementById('btn-pdf');
  if (wrap) wrap.classList.toggle('pdf-open', pdfOpen);
  if (btn)  btn.classList.toggle('active', pdfOpen);
  if (pdfOpen && pdfFile) loadPdf();
  else { const f = document.getElementById('pdf-frame'); if(f) f.src=''; }
}
function closePdf() {
  pdfOpen = false;
  const wrap  = document.getElementById('pdf-wrap');
  const frame = document.getElementById('pdf-frame');
  if (wrap)  wrap.classList.remove('pdf-open');
  if (frame) frame.src = '';
}

// ── Weryfikacja ───────────────────────────────
async function verify() {
  if (!current) return;
  const btn = document.getElementById('btn-v');
  if (!btn || btn.disabled) return;
  btn.disabled = true; btn.textContent = 'Przenoszenie…';
  const verifying = current;
  const r    = await fetch(`/api/verify/${verifying}`, {method:'POST'});
  const data = await r.json();
  const inv  = all.find(i => i.invoice_num === verifying);
  if (inv) inv.verified = true;
  verifyStack.push(verifying);
  updateProgress();
  showToast('Faktura zweryfikowana', current);
  if (data.next) { load(data.next); }
  else {
    current = null;
    document.getElementById('main').innerHTML =
      '<div class="empty-state" style="color:var(--green)">🎉 Wszystkie faktury zweryfikowane!</div>';
    closePdf();
  }
  renderList();
}

// ── Cofnij weryfikację (przycisk – cofa current) ──────────────────────────
async function unverify() {
  if (!current) return;
  await doUnverify(current);
}

// ── Cofnij ostatnią weryfikację (skrót U) ─────────────────────────────────
async function undoLastVerify() {
  if (!verifyStack.length) return;
  await doUnverify(verifyStack.pop());
}

async function doUnverify(num) {
  const r    = await fetch(`/api/unverify/${num}`, {method:'POST'});
  const data = await r.json();
  if (!data.ok) return;
  const inv = all.find(i => i.invoice_num === num);
  if (inv) inv.verified = false;
  updateProgress();
  showToast('Weryfikacja cofnięta', num);
  renderList();
  load(num);
}

// ── Odśwież ───────────────────────────────────
async function refreshInvoices() {
  const rr   = await fetch('/api/refresh');
  const info = await rr.json();
  if (info.extracted && info.extracted.length)
    showToast(`Wypakowano ${info.extracted.length} plik(i) ZIP`, info.extracted.join(', '));
  const r = await fetch('/api/invoices');
  all = await r.json();
  renderList();
  updateProgress();
}

// ── Toast ─────────────────────────────────────
function showToast(msg, val) {
  const old = document.getElementById('toast-el');
  if (old) old.remove();
  clearTimeout(toastTimer);
  const el = document.createElement('div');
  el.id = 'toast-el';
  el.className = 'toast';
  el.innerHTML = `<span class="toast-icon"></span><span>${msg}: </span><span class="mono" style="color:oklch(0.92 0.04 150)">${val}</span>`;
  document.body.appendChild(el);
  toastTimer = setTimeout(() => el.remove(), 2600);
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => showToast('Skopiowano WZ', text));
}

// ── Keyboard shortcuts ────────────────────────
document.addEventListener('keydown', e => {
  const tag = document.activeElement && document.activeElement.tagName;
  if (tag === 'INPUT' || tag === 'TEXTAREA') return;
  if (e.key === 'j' || e.key === 'ArrowUp')       { e.preventDefault(); navigateList(-1); }
  else if (e.key === 'k' || e.key === 'ArrowDown') { e.preventDefault(); navigateList(+1); }
  else if (e.key === 'm')  { e.preventDefault(); wzNav(+1); }
  else if (e.key === 'n')  { e.preventDefault(); wzNav(-1); }
  else if (e.key === ' ')  { e.preventDefault(); if (current) wzCheck(current); }
  else if (e.key === 'v')  { verify(); }
  else if (e.key === 'u')  { undoLastVerify(); }
  else if (e.key === 'p')  { if (current) togglePdf(); }
});

init();
</script>
</body>
</html>
"""

# ─── START ────────────────────────────────────────────────────────────────────

def find_free_port(start=5555):
    for p in range(start, start + 20):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(('', p))
            s.close()
            return p
        except OSError:
            pass
    return start

if __name__ == '__main__':
    port = find_free_port()
    url  = f'http://localhost:{port}'
    print(f'')
    print(f'  ╔══════════════════════════════════════╗')
    print(f'  ║   Faktury Żabka – weryfikacja        ║')
    print(f'  ╠══════════════════════════════════════╣')
    print(f'  ║  Adres:  {url:<28}  ║')
    print(f'  ║  Folder: {BASE_DIR[:28]:<28}  ║')
    print(f'  ║                                      ║')
    print(f'  ║  Zamknij to okno aby wyłączyć        ║')
    print(f'  ╚══════════════════════════════════════╝')
    print()
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)
