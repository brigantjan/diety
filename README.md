# SUPTel – Diéty a cestovné príkazy

Webová aplikácia (GitHub Pages), ktorá z mesačnej dochádzky vytvorí cestovné príkazy SR a ČR (diéty) a kontrolný prehľad.

- Prihlásenie: rovnaký e-mail a heslo ako Objednávky (Supabase projekt *dispecing-dev*).
- Nastavenia (zamestnanci, sadzby, časy) sú v Supabase v tabuľke `diety_nastavenia`. Každá zmena sa automaticky zálohuje do `diety_historia`.
- Výpočet beží priamo v prehliadači (Python cez Pyodide + openpyxl). Súbor dochádzky sa nikam neodosiela.
- Pred výpočtom sa skontroluje dochádzka (nadčas pri absencii, X na víkende, nevyplnené dni, sviatky v riadku 9).
- SR aj ČR príkaz majú na konci hárok **Súhrn** napojený vzorcami na hárky ľudí.
- Každý spracovaný mesiac sa uloží do archívu (`diety_archiv`); finálne upravené súbory majú prednosť.
- V repozitári **nie sú osobné údaje**. Šablóny obsahujú iba prázdne tlačivo VZOR.

## Súbory

| súbor | čo to je |
|---|---|
| `index.html`, `app.js` | stránka a jej logika |
| `config.js` | adresa Supabase a verzia výpočtového modulu |
| `py/engine.py` | výpočet diét (pravidlá SR/ČR, týždňovkári, farby, tlač) |
| `py/web_api.py` | prepojenie stránky s výpočtom |
| `sablony/` | prázdne tlačivá cestovného príkazu SR a ČR |
| `wheels/` | knižnica openpyxl pre Python v prehliadači |
| `sql/01_diety_tabulky.sql` | tabuľky a prístupové práva v Supabase (spúšťa sa raz) |
| `sql/02_diety_archiv.sql` | archív mesiacov – súčty po ľuďoch (spúšťa sa raz) |

## Zverejnenie na GitHub Pages

1. Na GitHube vytvor nový repozitár **diety** (Public, ako ostatné).
2. Klikni „uploading an existing file“ a pretiahni doň **celý obsah** tohto priečinka (aj podpriečinky). Potom daj „Commit changes“.
3. Otvor Settings → Pages → Source: *Deploy from a branch*, Branch: *main*, priečinok */ (root)* → Save.
4. Po 1–2 minútach beží aplikácia na `https://brigantjan.github.io/diety/`.

## Pridelenie prístupu ďalšiemu človeku

Supabase → SQL Editor:

```sql
insert into public.diety_pristup(user_id, poznamka)
select id, 'meno' from auth.users where email = 'email@suptel.sk'
on conflict do nothing;
```
