// =====================================================================
// KONFIGURÁCIA – Diéty a cestovné príkazy (SUPTel)
//
// Používa ten istý Supabase projekt ako Objednávky ("dispecing-dev"),
// takže prihlásenie je tým istým e-mailom a heslom.
//
// Kľúč je VEREJNÝ (publishable) – bez prihlásenia a bez pridelenia
// prístupu (tabuľka diety_pristup) sa cez neho nedá prečítať nič.
// =====================================================================

const SUPABASE_URL = "https://qpzktpuiyqwujbfhvxnb.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_B-2vQMUsXudJCUKSZFYTKg_RrXwf2_o";

// Výpočtový modul (Python v prehliadači). Verziu nemeň bez otestovania.
const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";

// TEST na DEV: kópia dát v schéme diety
const SUPABASE_SCHEMA = "diety";
