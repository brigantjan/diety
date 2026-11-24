// =====================================================================
// KONFIGURÁCIA – Diéty a cestovné príkazy (SUPTel)
//
// Používa ten istý Supabase projekt ako Objednávky (Dispečing PROD, schéma diety),
// takže prihlásenie je tým istým e-mailom a heslom.
//
// Kľúč je VEREJNÝ (publishable) – bez prihlásenia a bez pridelenia
// prístupu (tabuľka diety_pristup) sa cez neho nedá prečítať nič.
// =====================================================================

const SUPABASE_URL = "https://zdgfwuxfzhjldxnkhrzx.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_LaxKAWy-E-osNvXGJMfnjw_lco7NR2h";

// Výpočtový modul (Python v prehliadači). Verziu nemeň bez otestovania.
const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";

// Schéma v databáze PROD (oddelená od Dispečingu)
const SUPABASE_SCHEMA = "diety";
