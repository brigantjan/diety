# -*- coding: utf-8 -*-
"""
Generátor cestovných príkazov (diét) SR a ČR z mesačnej dochádzky – SUPTel s.r.o.

Použitie:  pretiahni súbor dochádzky na SPUSTI.bat  (alebo spusti a vyber súbor)
Výstup:    priečinok "Vystup" – SR príkaz, ČR príkaz, kontrolný prehľad
"""
import os, re, sys, calendar, datetime as dt, warnings
from copy import copy

warnings.filterwarnings('ignore')
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
NASTAVENIA = os.path.join(DATA, 'Nastavenia_diety.xlsx')
SABLONA_SR = os.path.join(DATA, 'sablony', 'sablona_SR.xlsx')
SABLONA_CZ = os.path.join(DATA, 'sablony', 'sablona_CZ.xlsx')
VYSTUP = os.path.join(HERE, 'Vystup')
from openpyxl.styles.colors import Color
FILL_ZLTA = PatternFill('solid', fgColor='FFFFFF00')
FILL_MODRA = PatternFill('solid', fgColor=Color(theme=5, tint=0.7999816888943144))
FILL_ZIADNA = PatternFill(fill_type=None)

MESIACE = ['január', 'február', 'marec', 'apríl', 'máj', 'jún', 'júl', 'august',
           'september', 'október', 'november', 'december']
ZNAME_KODY = {'X', 'D', 'SV', 'SC', 'SK', 'PZ', 'PNU', 'P', 'OCR', 'PO', 'A', 'Z', 'MD', 'RP'}
TYPY = {'dochádza', 'týždňovkár', 'týždňovkár nonstop', 'bez CP'}
FIRST_ROW = 7          # riadok 1. dňa v cestovnom príkaze (každý deň = 2 riadky)
WEEKDAY_TXT = {5: 'SO', 6: 'NE'}


# --------------------------------------------------------------------- pomocné
def clone_sheet(wb, src, title):
    """Kópia hárku vrátane nastavenia tlače (zalomenia, oblasť tlače, mierka)."""
    ws = wb.copy_worksheet(src)
    ws.title = title
    ws.col_breaks = copy(src.col_breaks)
    ws.row_breaks = copy(src.row_breaks)
    ws.sheet_view.zoomScale = src.sheet_view.zoomScale
    ws.sheet_view.showGridLines = src.sheet_view.showGridLines
    ws.oddHeader = copy(src.oddHeader)
    ws.oddFooter = copy(src.oddFooter)
    if src.print_area:
        area = src.print_area
        area = area.split('!')[-1] if isinstance(area, str) else area
        ws.print_area = area.replace('$', '')
    return ws


def fit_print(ws, last_row=85, brk_col=10, last_col=25):
    """Zníži mierku tlače, ak by sa hárok nezmestil na A4 (inak vznikne prázdna strana a rozhodí obojstrannú tlač)."""
    ps = ws.page_setup
    if ws.sheet_properties.pageSetUpPr is not None and ws.sheet_properties.pageSetUpPr.fitToPage:
        return
    m = ws.page_margins
    avail_h = (11.69 - (m.top or 0.75) - (m.bottom or 0.75)) * 72 - 4
    avail_w = (8.27 - (m.left or 0.7) - (m.right or 0.7)) * 72 - 4
    default_h = ws.sheet_format.defaultRowHeight or 15
    h = sum((ws.row_dimensions[r].height or default_h) for r in range(1, last_row + 1))

    def colw(c):
        d = ws.column_dimensions.get(openpyxl.utils.get_column_letter(c))
        w = d.width if d is not None and d.width else (ws.sheet_format.defaultColWidth or 9.14)
        return (w * 7 + 5) * 0.75
    w1 = sum(colw(c) for c in range(1, brk_col + 1))
    w2 = sum(colw(c) for c in range(brk_col + 1, last_col + 1))
    max_scale = int(min(avail_h / h, avail_w / max(w1, w2)) * 100)
    cur = ps.scale or 100
    if cur > max_scale:
        ps.scale = max(10, max_scale)


def setc(ws, ref, value, fmt=None):
    cell = ws[ref]
    if isinstance(cell, MergedCell):
        return
    cell.value = value
    if fmt:
        cell.number_format = fmt


def setv(ws, row, col, value, fmt=None):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        return
    cell.value = value
    if fmt:
        cell.number_format = fmt


def clear_range(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = None


def to_min(t):
    return t.hour * 60 + t.minute


def to_time(m):
    m = int(round(m))
    if m >= 24 * 60:
        return dt.time(23, 59)
    return dt.time(m // 60, m % 60)


def parse_time(v, default=None):
    if v is None or v == '':
        return default
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, dt.time):
        return v
    if isinstance(v, (int, float)):           # Excel zlomok dňa
        return to_time(float(v) * 24 * 60)
    m = re.match(r'^\s*(\d{1,2})[:.](\d{2})', str(v))
    if m:
        return dt.time(int(m.group(1)), int(m.group(2)))
    return default


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def safe_title(t):
    t = re.sub(r'[\[\]\:\*\?\/\\]', '', norm(t))[:31]
    return t or 'Hárok'


# ------------------------------------------------------------------ nastavenia
class Nastavenia:
    @classmethod
    def from_dict(cls, data):
        """Nastavenia z JSON (webová verzia – uložené v Supabase)."""
        self = cls.__new__(cls)
        self.path = None
        self.order = []
        self.zam = {}
        for z in data.get('zamestnanci', []):
            meno = norm(z.get('meno'))
            if not meno:
                continue
            self.zam[meno] = dict(
                meno=meno, typ=norm(z.get('typ')), harok=norm(z.get('harok')), bydlisko=norm(z.get('bydlisko')),
                t_odch=norm(z.get('t_odch')), t_pri=norm(z.get('t_pri')), miesto=norm(z.get('miesto')) or 'Žilina',
                ucel=norm(z.get('ucel')) or 'montážne práce',
                cz_samostatne='samostat' in norm(z.get('cz_mode')).lower())
        self.sadzby = sorted(
            (dt.date.fromisoformat(s['od']), float(s['a']), float(s['b']), float(s['c']), float(s['cz']))
            for s in data.get('sadzby', []) if s.get('od'))
        self.p = {norm(x.get('kluc')): x.get('hodnota') for x in data.get('parametre', []) if norm(x.get('kluc'))}
        self.vynimky = {}
        for v in data.get('vynimky', []):
            try:
                d = dt.date.fromisoformat(str(v.get('datum'))[:10])
            except ValueError:
                continue
            self.vynimky[(norm(v.get('meno')), d)] = dict(
                odch=parse_time(v.get('odch')), pri=parse_time(v.get('pri')),
                t_odch=norm(v.get('t_odch')), t_pri=norm(v.get('t_pri')))
        return self

    def __init__(self, path):
        self.path = path
        wb = openpyxl.load_workbook(path)
        self.zam = {}
        self.order = []
        z = wb['Zamestnanci']
        for r in range(2, z.max_row + 1):
            meno = norm(z.cell(r, 1).value)
            if not meno:
                continue
            self.zam[meno] = dict(
                row=r, meno=meno, typ=norm(z.cell(r, 2).value), harok=norm(z.cell(r, 3).value),
                bydlisko=norm(z.cell(r, 4).value), t_odch=norm(z.cell(r, 5).value),
                t_pri=norm(z.cell(r, 6).value), miesto=norm(z.cell(r, 7).value) or 'Žilina',
                ucel=norm(z.cell(r, 8).value) or 'montážne práce',
                cz_samostatne='samostat' in norm(z.cell(r, 10).value).lower())
        self.sadzby = []
        s = wb['Sadzby']
        for r in range(2, s.max_row + 1):
            d = s.cell(r, 1).value
            if isinstance(d, dt.datetime):
                self.sadzby.append((d.date(), float(s.cell(r, 2).value), float(s.cell(r, 3).value),
                                    float(s.cell(r, 4).value), float(s.cell(r, 5).value)))
        self.sadzby.sort()
        p = wb['Parametre']
        self.p = {}
        for r in range(2, p.max_row + 1):
            k = norm(p.cell(r, 1).value)
            if k:
                self.p[k] = p.cell(r, 2).value
        self.vynimky = {}
        v = wb['Výnimky']
        for r in range(2, v.max_row + 1):
            meno, d = norm(v.cell(r, 1).value), v.cell(r, 2).value
            if not meno or not isinstance(d, dt.datetime):
                continue
            self.vynimky[(meno, d.date())] = dict(
                odch=parse_time(v.cell(r, 3).value), pri=parse_time(v.cell(r, 4).value),
                t_odch=norm(v.cell(r, 5).value), t_pri=norm(v.cell(r, 6).value))

    def t(self, key, default):
        return parse_time(self.p.get(key), default)

    def s(self, key, default):
        v = norm(self.p.get(key))
        return v or default

    def sadzba(self, d):
        cur = None
        for row in self.sadzby:
            if row[0] <= d:
                cur = row
        if cur is None:
            raise ValueError(f'V nastaveniach (hárok Sadzby) chýba sadzba platná k {d:%d.%m.%Y}')
        return cur

    def pridaj_novych(self, mena):
        """Nových ľudí z dochádzky dopíše do nastavení (žlto) – treba im doplniť typ."""
        wb = openpyxl.load_workbook(self.path)
        z = wb['Zamestnanci']
        yellow = PatternFill('solid', fgColor='FFF2CC')
        used = {norm(z.cell(r, 3).value) for r in range(2, z.max_row + 1)}
        for meno in mena:
            title = meno.split(' ')[0]
            if title in used:
                title = f"{title} {meno.split(' ')[1][:1]}." if ' ' in meno else title + '2'
            used.add(title)
            r = z.max_row + 1
            vals = [meno, None, title, None, None, None, 'Žilina', 'montážne práce',
                    'NOVÝ – doplň typ, bydlisko a texty trasy']
            for c, val in enumerate(vals, 1):
                z.cell(r, c).value = val
                z.cell(r, c).fill = yellow
        wb.save(self.path)


# ------------------------------------------------------------------ dochádzka
class Dochadzka:
    def __init__(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb['Dochadzka'] if 'Dochadzka' in wb.sheetnames else wb.worksheets[0]
        self.ws = ws
        self.rok, self.mesiac = self._obdobie(ws)
        self.ndays = calendar.monthrange(self.rok, self.mesiac)[1]
        self.dates = [dt.date(self.rok, self.mesiac, d) for d in range(1, self.ndays + 1)]
        # hlavička: riadok s "Priezvisko a meno"
        hdr = None
        for r in range(1, 40):
            if 'priezvisko' in norm(ws.cell(r, 2).value).lower():
                hdr = r
                break
        if hdr is None:
            raise ValueError('V dochádzke som nenašiel riadok „Priezvisko a meno“.')
        self.col1 = 3   # stĺpec C = 1. deň
        # riadok pod hlavičkou: „Sv“ v deň sviatku
        self.sviatok = [norm(ws.cell(hdr + 1, self.col1 + i).value).upper() == 'SV' for i in range(self.ndays)]
        self.ludia = []  # (meno, [kódy])
        r = hdr + 2
        while r < hdr + 200:
            meno = norm(ws.cell(r, 2).value)
            if meno.lower().startswith('schválil'):
                break
            if meno:
                self.ludia.append((meno, r, [ws.cell(r, self.col1 + i).value for i in range(self.ndays)]))
            r += 1
        # tabuľka nadčasov
        self.nadcas = {}
        nr = None
        for rr in range(r, r + 60):
            if norm(ws.cell(rr, 2).value).lower().startswith('nadčas'):
                nr = rr
                break
        if nr is None:
            raise ValueError('V dochádzke som nenašiel tabuľku „Nadčas“.')
        rr = nr + 2
        for i, (meno, _, _) in enumerate(self.ludia):
            nm = norm(ws.cell(rr + i, 2).value)
            row = rr + i
            if nm and nm != meno and nm != '0':   # poradie nesedí – hľadaj podľa mena
                row = None
                for k in range(nr + 2, nr + 120):
                    if norm(ws.cell(k, 2).value) == meno:
                        row = k
                        break
            vals = []
            for d in range(self.ndays):
                v = ws.cell(row, self.col1 + d).value if row else None
                try:
                    vals.append(float(v) if v not in (None, '') else 0.0)
                except (TypeError, ValueError):
                    vals.append(0.0)
            self.nadcas[meno] = vals

    @staticmethod
    def _obdobie(ws):
        for r in range(1, 4):
            for c in range(1, 30):
                v = ws.cell(r, c).value
                if isinstance(v, dt.datetime):
                    return v.year, v.month
                s = norm(v).lower()
                m = re.search(r'mesiac\s+(\d{1,2})\s*r\.?\s*(\d{4})', s)
                if m:
                    return int(m.group(2)), int(m.group(1))
                for i, nm in enumerate(MESIACE):
                    m = re.search(nm + r'\s+(\d{4})', s)
                    if m:
                        return int(m.group(1)), i + 1
        raise ValueError('Z dochádzky sa nedá zistiť mesiac a rok (bunka B1 / O1).')


# ------------------------------------------------------------------ výpočet
def kod(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return ''
    if isinstance(v, (int, float)):
        return float(v)
    return v.strip().upper()


def sr_pasmo(hod, sad):
    if hod > 18:
        return sad[3]
    if hod > 12:
        return sad[2]
    if hod >= 5:
        return sad[1]
    return 0.0


def cz_suma(hod, plna):
    if hod > 12:
        return plna
    if hod > 6:
        return plna / 2
    if hod > 0:
        return plna / 4
    return 0.0


def vypocitaj(meno, z, kody, ot, doch, N, warn):
    """Vráti (sr_dni, cz_dni, cz_bloky). sr_dni[i] = dict(text=..) alebo dict(odch,pri,t_odch,t_pri)."""
    nd = doch.ndays
    dates = doch.dates
    kind = [None] * nd
    k = [kod(v) for v in kody]
    for i in range(nd):
        c, o = k[i], ot[i]
        if c == 'SC':
            kind[i] = 'CZ'
        elif c in ('X', 'SK'):
            kind[i] = 'SR'
        elif isinstance(c, float):
            kind[i] = 'SRH'                     # zapísané hodiny (dohoda)
        elif (c == '' or c == 'SV') and o > 0:
            kind[i] = 'OT'
        elif c:
            if c not in ZNAME_KODY:
                warn(f'{meno}: neznámy kód „{kody[i]}“ dňa {dates[i]:%d.%m.} – zapísaný ako text')
            kind[i] = 'TXT'
        else:
            kind[i] = 'EMPTY'
    # deň len s nadčasom (víkend/sviatok) patrí tam, kde bol naposledy v práci
    for i in range(nd):
        if kind[i] != 'OT':
            continue
        cntry = None
        for j in list(range(i - 1, -1, -1)) + list(range(i + 1, nd)):
            if kind[j] in ('CZ', 'SR', 'SRH'):
                cntry = 'CZ' if kind[j] == 'CZ' else 'SR'
                break
        kind[i] = 'CZ' if cntry == 'CZ' else ('SR_OT' if ot[i] >= 5 else 'TXT')

    def txt(i):
        c = k[i]
        if isinstance(c, str) and c:
            return c
        if dates[i].weekday() >= 5:
            return WEEKDAY_TXT[dates[i].weekday()]
        return 'SV' if doch.sviatok[i] else None

    sr = [None] * nd
    typ = z['typ']
    miesto = z['miesto']
    pokracuje_zac = dates[0].weekday() != 0          # mesiac nezačína pondelkom
    pokracuje_kon = dates[-1].weekday() <= 3          # mesiac končí Po–Št

    if typ == 'dochádza':
        o1 = N.t('Dochádza – odchod', dt.time(7, 0))
        p1 = N.t('Dochádza – príchod bez nadčasu', dt.time(15, 30))
        we = N.t('Víkend/sviatok – odchod', dt.time(7, 0))
        for i in range(nd):
            if kind[i] == 'SR':
                sr[i] = dict(odch=o1, pri=to_time(to_min(p1) + ot[i] * 60))
            elif kind[i] == 'SRH':
                sr[i] = dict(odch=o1, pri=to_time(to_min(o1) + (k[i] + ot[i]) * 60))
            elif kind[i] == 'SR_OT':
                sr[i] = dict(odch=we, pri=to_time(to_min(we) + ot[i] * 60))
            if sr[i]:
                sr[i].update(t_odch=z['t_odch'] or f"{z['bydlisko']}-{miesto}",
                             t_pri=z['t_pri'] or f"{miesto}-{z['bydlisko']}")
    else:   # týždňovkári
        nonstop = typ == 'týždňovkár nonstop'
        worked = [kind[i] in ('SR', 'SRH', 'SR_OT') for i in range(nd)]
        doma = z['t_odch'] or z['bydlisko']
        i = 0
        while i < nd:
            if not worked[i]:
                i += 1
                continue
            s = i
            while i + 1 < nd and worked[i + 1]:
                i += 1
            e = i
            for d in range(s, e + 1):
                first = d == s and not (s == 0 and pokracuje_zac)
                last = d == e and not (e == nd - 1 and pokracuje_kon)
                if nonstop:
                    sr[d] = dict(odch=N.t('Nonstop – odchod', dt.time(0, 0)),
                                 pri=N.t('Nonstop – príchod', dt.time(23, 59)), t_odch=miesto, t_pri=miesto)
                    continue
                odch = N.t('Týždňovkár – odchod', dt.time(5, 0))
                pri = N.t('Týždňovkár – príchod (posledný deň)', dt.time(19, 0)) if last \
                    else N.t('Týždňovkár – príchod (bežný deň)', dt.time(23, 1))
                sr[d] = dict(odch=odch, pri=pri, t_odch=doma if first else miesto,
                             t_pri=(z['t_pri'] or doma) if last else miesto)
            i += 1

    # výnimky
    for i in range(nd):
        v = N.vynimky.get((meno, dates[i]))
        if not v:
            continue
        base = sr[i] or dict(odch=N.t('Dochádza – odchod', dt.time(7, 0)),
                             pri=to_time(to_min(N.t('Dochádza – príchod bez nadčasu', dt.time(15, 30))) + ot[i] * 60),
                             t_odch=z['t_odch'], t_pri=z['t_pri'])
        if v['odch']:
            base['odch'] = v['odch']
        if v['pri']:
            base['pri'] = v['pri']
        if v['t_odch']:
            base['t_odch'] = v['t_odch']
        if v['t_pri']:
            base['t_pri'] = v['t_pri']
        sr[i] = base
        if kind[i] == 'CZ':
            warn(f'{meno}: výnimka na {dates[i]:%d.%m.} je v deň služobnej cesty v ČR – v SR príkaze prepísala SC')
            kind[i] = 'SR'

    for i in range(nd):
        if sr[i] is None:
            sr[i] = dict(text='SC' if kind[i] == 'CZ' else txt(i))
        sr[i]['volno'] = dates[i].weekday() >= 5 or doch.sviatok[i] or k[i] == 'SV'
        sr[i]['cz'] = kind[i] == 'CZ' and 'text' in sr[i]

    # ---- ČR
    cz = [None] * nd
    bloky = []
    border = N.s('ČR – hraničný prechod', 'Svrčinovec')
    site = N.s('ČR – miesto výkonu', 'Bruntál')
    i = 0
    while i < nd:
        if kind[i] != 'CZ':
            i += 1
            continue
        s = i
        while i + 1 < nd and kind[i + 1] == 'CZ':
            i += 1
        e = i
        if s != e and z.get('cz_samostatne'):
            e = s
            i = s
        if s == e:
            cz[s] = dict(single=True, odch=N.t('ČR jednodňová – odchod', dt.time(7, 0)),
                         pri=N.t('ČR jednodňová – príchod', dt.time(16, 30)),
                         t_odch=f'{border} -{site}', t_pri=f'{site} - {border}')
            bloky.append((dates[s], dates[e], border, border))
        else:
            cs = s == 0 and pokracuje_zac
            ce = e == nd - 1 and pokracuje_kon
            for d in range(s, e + 1):
                if d == s and not cs:
                    cz[d] = dict(odch=N.t('ČR viacdňová – prvý deň odchod', dt.time(9, 9)),
                                 pri=N.t('ČR viacdňová – prvý deň príchod', dt.time(21, 15)), t_odch=border, t_pri=site)
                elif d == e and not ce:
                    cz[d] = dict(odch=N.t('ČR viacdňová – posledný deň odchod', dt.time(0, 0)),
                                 pri=N.t('ČR viacdňová – posledný deň príchod', dt.time(16, 16)), t_odch=site, t_pri=border)
                else:
                    cz[d] = dict(odch=N.t('ČR viacdňová – stredný deň odchod', dt.time(0, 0)),
                                 pri=N.t('ČR viacdňová – stredný deň príchod', dt.time(19, 1)), t_odch=site, t_pri=site)
                cz[d]['single'] = False
            bloky.append((dates[s], dates[e], site if cs else border, site if ce else border))
        i += 1
    for i in range(nd):
        if cz[i] is None:
            cz[i] = dict(text=WEEKDAY_TXT.get(dates[i].weekday()))
        else:
            cz[i]['cz'] = True
        cz[i]['volno'] = sr[i]['volno']
    return sr, cz, bloky


# ------------------------------------------------------------------ zápis
def priprav_harok(wb, meno, title):
    """Nájde hárok osoby v šablóne (podľa mena v D3), inak skopíruje VZOR."""
    for ws in wb.worksheets:
        if ws.title not in ('VZOR',) and 'Pozn' not in ws.title and norm(ws['D3'].value) == meno:
            ws.title = '_tmp_' + str(id(ws))[-8:]
            ws._target_title = title
            return ws
    ws = clone_sheet(wb, wb['VZOR'], '_tmp_' + str(len(wb.worksheets)))
    ws._target_title = title
    return ws


def vyfarbi(ws, r, c1, c2, fill):
    for rr in (r, r + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(rr, c)
            if not isinstance(cell, MergedCell):
                cell.fill = fill


def farba_dna(ws, r, d):
    """Žltá K = víkend/sviatok, žltý celý riadok = odpracovaný víkend/sviatok, modrá L–Y = deň v ČR."""
    cesta = 'text' not in d
    if d.get('cz'):
        vyfarbi(ws, r, 12, 25, FILL_MODRA)
        if d.get('volno'):
            vyfarbi(ws, r, 11, 11, FILL_ZLTA)
    elif d.get('volno') and cesta:
        vyfarbi(ws, r, 11, 25, FILL_ZLTA)
    elif d.get('volno'):
        vyfarbi(ws, r, 11, 11, FILL_ZLTA)


def zapis_dni(ws, doch, dni, sumy_fn, fmt_suma, aus):
    clear_range(ws, FIRST_ROW, FIRST_ROW + 61, 11, 25)
    vyfarbi(ws, FIRST_ROW, 11, 25, FILL_ZIADNA)
    for i in range(31):
        r = FIRST_ROW + 2 * i
        vyfarbi(ws, r, 11, 25, FILL_ZIADNA)
        if i >= doch.ndays:
            continue
        farba_dna(ws, r, dni[i])
        setv(ws, r, 11, dt.datetime.combine(doch.dates[i], dt.time()), 'd/m;@')
        d = dni[i]
        if 'text' in d:
            if d['text']:
                setv(ws, r, 20, d['text'])
                setv(ws, r, 24, f'=T{r}')
            continue
        setv(ws, r, 12, 'Odchod ')
        setv(ws, r + 1, 12, 'Príchod')
        setv(ws, r, 13, d.get('t_odch') or None)
        setv(ws, r + 1, 13, d.get('t_pri') or None)
        setv(ws, r, 14, d['odch'], 'hh:mm')
        setv(ws, r + 1, 14, d['pri'], 'hh:mm')
        setv(ws, r, 15, aus)
        setv(ws, r, 17, f'=N{r + 1}-N{r}', 'hh:mm')
        setv(ws, r, 20, sumy_fn(r, d), fmt_suma)
        setv(ws, r, 24, f'=T{r}', fmt_suma)


def hodiny(d):
    return max(0, to_min(d['pri']) - to_min(d['odch'])) / 60


# ------------------------------------------------------------------ sviatky a kontrola dochádzky
def _velka_noc(rok):
    a, b, c = rok % 19, rok // 100, rok % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    den = ((h + l - 7 * m + 114) % 31) + 1
    return dt.date(rok, mes, den)


def dni_pracovneho_pokoja(rok):
    """Dni pracovného pokoja v SR (zákon 241/1993 v znení konsolidačných zmien)."""
    vn = _velka_noc(rok)
    d = {
        dt.date(rok, 1, 1): 'Deň vzniku SR', dt.date(rok, 1, 6): 'Zjavenie Pána',
        vn - dt.timedelta(days=2): 'Veľký piatok', vn + dt.timedelta(days=1): 'Veľkonočný pondelok',
        dt.date(rok, 5, 1): 'Sviatok práce', dt.date(rok, 7, 5): 'Sviatok sv. Cyrila a Metoda',
        dt.date(rok, 8, 29): 'Výročie SNP', dt.date(rok, 11, 1): 'Sviatok všetkých svätých',
        dt.date(rok, 12, 24): 'Štedrý deň', dt.date(rok, 12, 25): 'Prvý sviatok vianočný',
        dt.date(rok, 12, 26): 'Druhý sviatok vianočný',
    }
    if rok != 2026:                       # v roku 2026 dočasne nie sú dňami pracovného pokoja
        d[dt.date(rok, 5, 8)] = 'Deň víťazstva nad fašizmom'
        d[dt.date(rok, 9, 15)] = 'Sedembolestná Panna Mária'
    if rok <= 2023:
        d[dt.date(rok, 9, 1)] = 'Deň Ústavy SR'
    if rok <= 2024:
        d[dt.date(rok, 11, 17)] = 'Deň boja za slobodu a demokraciu'
    if rok <= 2025:
        d[dt.date(rok, 10, 28)] = 'Deň vzniku česko-slovenského štátu'
    return d


ABSENCIE = {'D', 'P', 'PNU', 'OCR', 'PO', 'A', 'Z', 'MD', 'RP', 'PZ'}
DNI_TYZDNA = ['pondelok', 'utorok', 'streda', 'štvrtok', 'piatok', 'sobota', 'nedeľa']


def _rozsahy(dni):
    """[3,4,5,9] -> '3.–5., 9.'"""
    out, i = [], 0
    while i < len(dni):
        j = i
        while j + 1 < len(dni) and dni[j + 1] == dni[j] + 1:
            j += 1
        out.append(f'{dni[i]}.' if i == j else f'{dni[i]}.–{dni[j]}.')
        i = j + 1
    return ', '.join(out)


def kontrola_dochadzky(doch, N):
    """Podozrivé miesta v dochádzke. Vracia zoznam dict(meno, text, druh) – druh: 'chyba' / 'info'."""
    out = []
    sviatky = dni_pracovneho_pokoja(doch.rok)
    m = doch.mesiac
    for i, d in enumerate(doch.dates):
        if d.weekday() >= 5:
            continue
        if d in sviatky and not doch.sviatok[i]:
            out.append(dict(meno='', druh='chyba', text=f'{d.day}.{m}. je deň pracovného pokoja ({sviatky[d]}), '
                                                         f'ale v riadku 9 dochádzky nie je „Sv“.'))
        if doch.sviatok[i] and d not in sviatky:
            out.append(dict(meno='', druh='chyba', text=f'{d.day}.{m}. je v riadku 9 označený ako sviatok („Sv“), '
                                                         f'ale podľa zákona to nie je deň pracovného pokoja.'))
    for meno, _, kody in doch.ludia:
        z = N.zam.get(meno)
        if z is None or z['typ'] == 'bez CP':
            continue
        ot = doch.nadcas.get(meno, [0.0] * doch.ndays)
        nevyplnene = []
        for i, d in enumerate(doch.dates):
            c, o = kod(kody[i]), ot[i]
            volno = d.weekday() >= 5 or doch.sviatok[i] or d in sviatky
            dtxt = f'{d.day}.{m}. ({DNI_TYZDNA[d.weekday()]})'
            h = f'{o:g}'.replace('.', ',')
            if isinstance(c, str) and c in ABSENCIE and o > 0:
                out.append(dict(meno=meno, druh='chyba', text=f'{dtxt}: nadčas {h} h, ale v dochádzke je „{c}“.'))
            elif c in ('X', 'SK') and d.weekday() >= 5:
                out.append(dict(meno=meno, druh='chyba', text=f'{dtxt}: zapísané „{c}“ na víkende – počíta sa ako celá zmena '
                                                              f'(7:00–15:30 + nadčas). Ak robil len nadčas, nechaj deň prázdny.'))
            elif c == '' and not volno:
                if o > 0:
                    out.append(dict(meno=meno, druh='chyba', text=f'{dtxt}: nadčas {h} h, ale chýba kód (X, SC…) – '
                                                                  f'počíta sa len ako {h} h od 7:00.'))
                else:
                    nevyplnene.append(d.day)
            if o > 0 and ((c in ('X', 'SK', 'SC') and o > 12) or (c in ('', 'SV') and o > 16)):
                out.append(dict(meno=meno, druh='chyba', text=f'{dtxt}: neobvykle vysoký nadčas {h} h – skontroluj preklep.'))
        if nevyplnene:
            out.append(dict(meno=meno, druh='info', text=f'nevyplnené pracovné dni: {_rozsahy(nevyplnene)}'))
    return out


def _q(title):
    return "'" + title.replace("'", "''") + "'"


def pridaj_suhrn(wb, used, druh, doch, N):
    """Súhrnný hárok na konci zošita – vzorce ťahajú hodnoty z hárkov ľudí (po ručnej úprave sa prepočíta)."""
    from openpyxl.styles import Border, Side
    ws = wb.create_sheet('Súhrn')
    sad = N.sadzba(doch.dates[0])
    mena = 'Kč' if druh == 'CZ' else '€'
    if druh == 'SR':
        pasma = [(sad[1], f'{sad[1]:.2f} €'.replace('.', ',')), (sad[2], f'{sad[2]:.2f} €'.replace('.', ',')),
                 (sad[3], f'{sad[3]:.2f} €'.replace('.', ','))]
        nadpis = 'Súhrn diét – tuzemské cesty (SR)'
        dni_txt = 'Dni s diétou'
    else:
        pasma = [(sad[4] / 4, f'{sad[4] / 4:g} Kč'), (sad[4] / 2, f'{sad[4] / 2:g} Kč'), (sad[4], f'{sad[4]:g} Kč')]
        nadpis = 'Súhrn diét – zahraničné cesty (ČR)'
        dni_txt = 'Dni v ČR'
    ws['A1'] = f'{nadpis} – {MESIACE[doch.mesiac - 1]} {doch.rok}'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = f'Obdobie: {doch.mesiac:02d}/{doch.rok} | {druh}'
    ws['A2'].font = Font(size=9, color='808080')
    hdr = ['Meno', dni_txt, *[p[1] for p in pasma], f'Spolu {mena}']
    thin = Side(style='thin', color='B7C3D0')
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, v in enumerate(hdr, 1):
        cell = ws.cell(4, c, v)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = brd
    ws.cell(3, 3, 'Počet dní v pásme').font = Font(italic=True, size=9, color='606060')
    fmt = '#,##0.00" €"' if druh == 'SR' else '#,##0\\ "Kč"'
    r = 5
    for sh in used:
        ref = _q(sh.title)
        rng = f'{ref}!$T$7:$T$68'
        ws.cell(r, 1, norm(sh['D3'].value) or sh.title)
        ws.cell(r, 1).hyperlink = f"#{ref}!A1"
        ws.cell(r, 1).font = Font(color='1F4E78', underline='single')
        ws.cell(r, 2, f'=COUNTIF({rng},">0")')
        for j, (val, _) in enumerate(pasma):
            ws.cell(r, 3 + j, f'=COUNTIF({rng},{val:g})')
        ws.cell(r, 6, f'={ref}!X79').number_format = fmt
        for c in range(1, 7):
            ws.cell(r, c).border = brd
            if c > 1:
                ws.cell(r, c).alignment = Alignment(horizontal='right' if c == 6 else 'center')
        if (r - 5) % 2:
            for c in range(1, 7):
                ws.cell(r, c).fill = PatternFill('solid', fgColor='F2F6FA')
        r += 1
    ws.cell(r, 1, 'SPOLU').font = Font(bold=True)
    for c in range(2, 7):
        L = openpyxl.utils.get_column_letter(c)
        cell = ws.cell(r, c, f'=SUM({L}5:{L}{r - 1})')
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='right' if c == 6 else 'center')
    for c in range(1, 7):
        ws.cell(r, c).border = Border(top=Side(style='medium', color='1F4E78'), bottom=thin, left=thin, right=thin)
        ws.cell(r, c).fill = PatternFill('solid', fgColor='E3EDF9')
    ws.cell(r, 6).number_format = fmt
    note = ws.cell(r + 2, 1, 'Hodnoty sa berú priamo z hárkov jednotlivých ľudí (stĺpec Stravné a riadok Doplatok). '
                             'Po ručnej úprave času alebo sumy sa súhrn prepočíta sám.')
    note.font = Font(size=9, italic=True, color='606060')
    note.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=6)
    ws.row_dimensions[r + 2].height = 26
    for c, w in zip('ABCDEF', (30, 13, 12, 12, 12, 16)):
        ws.column_dimensions[c].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = 'A5'
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.6
    ws.page_margins.top = ws.page_margins.bottom = 0.7
    ws.print_area = f'A1:F{r + 2}'
    return ws


def nacitaj_suhrn(path):
    """Prečíta hárok Súhrn z finálneho (v Exceli uloženého) súboru – pre archív."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if 'Súhrn' not in wb.sheetnames:
        raise ValueError('V súbore nie je hárok „Súhrn“ – nahraj SR alebo ČR príkaz vytvorený touto aplikáciou.')
    ws = wb['Súhrn']
    rows = list(ws.iter_rows(min_row=1, max_row=200, max_col=6, values_only=True))
    m = re.search(r'Obdobie:\s*(\d{2})/(\d{4})\s*\|\s*(SR|CZ)', str(rows[1][0] or ''))
    if not m:
        raise ValueError('Hárok „Súhrn“ nemá údaj o období – súbor nie je z tejto aplikácie.')
    ludia = []
    for row in rows[4:]:
        meno = norm(row[0])
        if not meno or meno == 'SPOLU':
            if meno == 'SPOLU':
                break
            continue
        if row[5] is None or row[1] is None:
            raise ValueError('Súbor nemá vypočítané hodnoty – otvor ho v Exceli, ulož (Ctrl+S) a nahraj znova.')
        ludia.append(dict(meno=meno, dni=int(row[1] or 0), spolu=round(float(row[5] or 0), 2)))
    return dict(mesiac=int(m.group(1)), rok=int(m.group(2)), druh=m.group(3), ludia=ludia)


def main(doch_path, out_dir=VYSTUP, N=None):
    os.makedirs(out_dir, exist_ok=True)
    if N is None:
        N = Nastavenia(NASTAVENIA)
    doch = Dochadzka(doch_path)
    upoz = []
    warn = upoz.append
    mes_txt = f'{doch.mesiac:02d}_{doch.rok}'
    print(f'Dochádzka: {MESIACE[doch.mesiac - 1]} {doch.rok}, {len(doch.ludia)} ľudí')

    for m, _, _ in doch.ludia:
        if m not in N.zam:
            warn(f'{m}: nie je v zozname zamestnancov – cestovný príkaz sa nevytvoril')

    sr_wb = openpyxl.load_workbook(SABLONA_SR)
    cz_wb = openpyxl.load_workbook(SABLONA_CZ)
    aus = N.s('Dopravný prostriedok', 'AUS')
    stredisko = N.s('Stredisko (do názvu súboru)', 'Žilina')
    schvaluje = N.s('So spôsobom vykonania súhlasí', '') or None
    first, last = doch.dates[0], doch.dates[-1]
    d_first = dt.datetime.combine(first, dt.time())
    d_last = dt.datetime.combine(last, dt.time())
    sr_used, cz_used, prehlad = [], [], []
    kontrola = kontrola_dochadzky(doch, N)
    used_titles = set()

    for meno, _, kody in doch.ludia:
        z = N.zam.get(meno)
        if z is None:
            continue
        if z['typ'] not in TYPY:
            warn(f'{meno}: v nastaveniach chýba alebo je zlý typ („{z["typ"]}“) – preskočený')
            continue
        if z['typ'] == 'bez CP':
            continue
        if not z['bydlisko']:
            warn(f'{meno}: v nastaveniach chýba bydlisko')
        title = safe_title(z['harok'] or meno.split(' ')[0])
        base, n = title, 2
        while title in used_titles:
            title = f'{base[:28]} {n}'
            n += 1
        used_titles.add(title)
        ot = doch.nadcas.get(meno, [0.0] * doch.ndays)
        sr, cz, bloky = vypocitaj(meno, z, kody, ot, doch, N, warn)

        # ---------- SR
        ws = priprav_harok(sr_wb, meno, title)
        setc(ws, 'D3', meno)
        setc(ws, 'C4', z['bydlisko'])
        setc(ws, 'I1', None)
        setc(ws, 'E69', schvaluje)
        setc(ws, 'B9', d_first, 'd/mmmm\\ yyyy')
        setc(ws, 'I9', d_last, 'd/mmmm\\ yyyy')
        setc(ws, 'D9', z['miesto'])
        setc(ws, 'G9', z['ucel'])
        setc(ws, 'B10', '=C4')
        setc(ws, 'I10', '=C4')

        def sr_suma(r, d):
            sad = N.sadzba(doch.dates[(r - FIRST_ROW) // 2])
            h = f'ROUND((N{r + 1}-N{r})*24,2)'
            return (f'=IF(OR(N{r}="",N{r + 1}=""),"",IF({h}>18,{sad[3]},IF({h}>12,{sad[2]},'
                    f'IF({h}>=5,{sad[1]},0))))')
        zapis_dni(ws, doch, sr, sr_suma, '#,##0.00" €"', aus)
        sr_used.append(ws)

        sr_sum, pasma = 0.0, {}
        for i, d in enumerate(sr):
            if 'text' in d:
                continue
            h = hodiny(d)
            val = sr_pasmo(h, N.sadzba(doch.dates[i]))
            sr_sum += val
            pasma[val] = pasma.get(val, 0) + 1
            if val == 0:
                warn(f'{meno}: {doch.dates[i]:%d.%m.} cesta {h:.1f} h – pod 5 h, diéta 0 €')

        # ---------- ČR
        cz_sum, cz_dni = 0.0, 0
        if bloky:
            ws = priprav_harok(cz_wb, meno, title)
            setc(ws, 'D3', meno)
            setc(ws, 'C4', z['bydlisko'])
            setc(ws, 'I1', None)
            setc(ws, 'E69', schvaluje)
            clear_range(ws, 9, 26, 2, 10)
            setc(ws, 'G36', d_last, 'd/mmmm\\ yyyy')
            miesto_cz = N.s('ČR – miesto rokovania', 'CZ - Bruntál')
            ucel_cz = N.s('ČR – účel cesty', 'Zabezpečenie prác')
            slots = [9, 12, 15, 18, 21]
            if len(bloky) > len(slots):
                warn(f'{meno}: {len(bloky)} ciest do ČR, formulár má miesto na {len(slots)} – '
                     f'v hlavičke je posledná cesta zlúčená ({bloky[len(slots) - 1][0]:%d.%m.}–{bloky[-1][1]:%d.%m.})')
                bloky = bloky[:len(slots) - 1] + [(bloky[len(slots) - 1][0], bloky[-1][1], bloky[len(slots) - 1][2], bloky[-1][3])]
            for (b_s, b_e, p_s, p_e), r in zip(bloky, slots):
                setv(ws, r, 2, dt.datetime.combine(b_s, dt.time()), 'd/mmmm\\ yyyy')
                setv(ws, r, 4, miesto_cz)
                setv(ws, r, 7, ucel_cz)
                setv(ws, r, 9, dt.datetime.combine(b_e, dt.time()), 'd/mmmm\\ yyyy')
                setv(ws, r + 1, 2, p_s)
                setv(ws, r + 1, 9, p_e)

            def cz_suma_fn(r, d):
                plna = N.sadzba(doch.dates[(r - FIRST_ROW) // 2])[4]
                if not d['single']:
                    return plna
                h = f'ROUND((N{r + 1}-N{r})*24,2)'
                return (f'=IF(OR(N{r}="",N{r + 1}=""),"",IF({h}>12,{plna},IF({h}>6,{plna / 2},'
                        f'IF({h}>0,{plna / 4},0))))')
            zapis_dni(ws, doch, cz, cz_suma_fn, '#,##0.00\\ [$CZK]', aus)
            cz_used.append(ws)
            for i, d in enumerate(cz):
                if 'text' in d:
                    continue
                cz_dni += 1
                plna = N.sadzba(doch.dates[i])[4]
                cz_sum += cz_suma(hodiny(d), plna) if d['single'] else plna
        prehlad.append([meno, z['typ'], title, sum(pasma.values()),
                        *[pasma.get(N.sadzba(first)[j], 0) for j in (1, 2, 3)],
                        round(sr_sum, 2), cz_dni, round(cz_sum, 2), len(bloky)])

    # odstráň nepoužité hárky, zoraď podľa dochádzky
    def finalize(wb, used):
        for ws in list(wb.worksheets):
            if 'Pozn' in ws.title:
                continue
            if ws not in used:
                wb.remove(ws)
        for ws in used:
            ws.title = ws._target_title
            fit_print(ws)
        pozn = [w for w in wb.worksheets if 'Pozn' in w.title]
        wb._sheets = pozn + used
        act = len(pozn) if used else 0
        wb.active = act
        for idx, w in enumerate(wb.worksheets):
            w.sheet_view.tabSelected = (idx == act)
        if 'Excel_BuiltIn_Print_Area_1' in wb.defined_names:
            del wb.defined_names['Excel_BuiltIn_Print_Area_1']

    finalize(sr_wb, sr_used)
    if sr_used:
        pridaj_suhrn(sr_wb, sr_used, 'SR', doch, N)
    out_sr = os.path.join(out_dir, f'SR_Cestovný príkaz_{mes_txt}_{stredisko}.xlsx')
    out_cz = os.path.join(out_dir, f'CR_Cestovný príkaz_{mes_txt}_{stredisko}.xlsx')
    out_k = os.path.join(out_dir, f'Kontrola_diet_{mes_txt}.xlsx')
    saved = []
    for wb, path in ((sr_wb, out_sr), (cz_wb, out_cz)):
        if wb is cz_wb:
            finalize(cz_wb, cz_used)
            if not cz_used:
                print('ČR: tento mesiac nikto nebol v ČR – súbor sa nevytvára.')
                continue
            pridaj_suhrn(cz_wb, cz_used, 'CZ', doch, N)
        try:
            wb.save(path)
            saved.append(path)
        except PermissionError:
            raise SystemExit(f'Súbor {os.path.basename(path)} je otvorený v Exceli – zatvor ho a spusti znova.')

    # kontrolný prehľad
    k = openpyxl.Workbook()
    s = k.active
    s.title = 'Prehľad'
    sad = N.sadzba(first)
    s.append([f'Kontrola diét – {MESIACE[doch.mesiac - 1]} {doch.rok}'])
    s.append(['Meno', 'Typ', 'Hárok', 'SR dni s diétou', f'{sad[1]:.2f} €', f'{sad[2]:.2f} €', f'{sad[3]:.2f} €',
              'SR spolu €', 'ČR dni', 'ČR spolu Kč', 'ČR ciest'])
    for row in prehlad:
        s.append(row)
    tr = s.max_row + 1
    s.cell(tr, 1, 'SPOLU').font = Font(bold=True)
    for c in (4, 5, 6, 7, 8, 9, 10, 11):
        L = openpyxl.utils.get_column_letter(c)
        s.cell(tr, c, f'=SUM({L}3:{L}{tr - 1})').font = Font(bold=True)
    s['A1'].font = Font(bold=True, size=14)
    for c in s[2]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F4E78')
        c.alignment = Alignment(wrap_text=True, horizontal='center')
    for i, w in enumerate([24, 18, 12, 10, 9, 9, 9, 12, 8, 12, 8], 1):
        s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for r in range(3, tr + 1):
        s.cell(r, 8).number_format = '#,##0.00'
        s.cell(r, 10).number_format = '#,##0'
    s.cell(tr + 2, 1, 'ČR jednodňové cesty majú predvyplnené časy – po oprave podľa GPS sa suma v ČR príkaze prepočíta sama.')
    kd = k.create_sheet('Kontrola dochádzky')
    kd.append(['Meno', 'Zistenie'])
    for c in kd[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F4E78')
    for x in kontrola or [dict(meno='', text='Bez zistení.', druh='info')]:
        kd.append([x['meno'] or 'celá dochádzka', x['text']])
        if x['druh'] == 'chyba':
            kd.cell(kd.max_row, 2).font = Font(color='9C2A00')
    kd.column_dimensions['A'].width = 26
    kd.column_dimensions['B'].width = 110
    u = k.create_sheet('Upozornenia')
    u.append(['Upozornenie'])
    u['A1'].font = Font(bold=True)
    for w in upoz or ['Bez upozornení.']:
        u.append([w])
    u.column_dimensions['A'].width = 120
    try:
        k.save(out_k)
        saved.append(out_k)
    except PermissionError:
        print('Kontrolný súbor je otvorený – neuložený.')

    print()
    for p in saved:
        print('  ✔', os.path.basename(p))
    if upoz:
        print('\nUPOZORNENIA:')
        for w in upoz:
            print('  •', w)
    return dict(rok=doch.rok, mesiac=doch.mesiac, mesiac_txt=f'{MESIACE[doch.mesiac - 1]} {doch.rok}',
                files=saved, warnings=upoz, prehlad=prehlad, kontrola=kontrola,
                sadzby=[N.sadzba(first)[j] for j in (1, 2, 3)])


if __name__ == '__main__':
    res = main(sys.argv[1])
    print(res['files'])
    for w in res['warnings']:
        print(' •', w)
