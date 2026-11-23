# -*- coding: utf-8 -*-
"""Rozhranie medzi webovou stránkou (JavaScript) a výpočtom diét (engine.py) – beží v prehliadači cez Pyodide."""
import json, os, shutil
import engine

TYPY = sorted(engine.TYPY)


def _stav(z):
    if z is None:
        return 'novy'
    if z.get('typ') not in engine.TYPY:
        return 'neuplny'
    if z.get('typ') != 'bez CP' and not engine.norm(z.get('bydlisko')):
        return 'neuplny'
    return 'ok'


def analyzuj(doch_path, settings_json):
    data = json.loads(settings_json)
    zam = {engine.norm(z.get('meno')): z for z in data.get('zamestnanci', [])}
    d = engine.Dochadzka(doch_path)
    ludia = [dict(meno=m, stav=_stav(zam.get(m))) for m, _, _ in d.ludia]
    return json.dumps(dict(rok=d.rok, mesiac=d.mesiac, mesiac_txt=f'{engine.MESIACE[d.mesiac - 1]} {d.rok}',
                           ludia=ludia), ensure_ascii=False)


def generuj(doch_path, settings_json, out_dir='/tmp/vystup'):
    shutil.rmtree(out_dir, ignore_errors=True)
    N = engine.Nastavenia.from_dict(json.loads(settings_json))
    res = engine.main(doch_path, out_dir, N)
    res['files'] = [os.path.basename(p) for p in res['files']]
    res['out_dir'] = out_dir
    return json.dumps(res, ensure_ascii=False, default=str)


def zabal(out_dir, nazov):
    """Všetky výstupné súbory do jedného ZIP-u."""
    import zipfile
    path = os.path.join('/tmp', nazov)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in sorted(os.listdir(out_dir)):
            z.write(os.path.join(out_dir, n), n)
    return path


def nacitaj_suhrn(path):
    """Súhrn z finálneho SR/ČR súboru (uloženého v Exceli) – pre archív."""
    return json.dumps(engine.nacitaj_suhrn(path), ensure_ascii=False)
