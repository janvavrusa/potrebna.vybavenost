"use strict";

const REQUIRED_SHEETS = ["Model_Vybavenosti", "Pravidla_a_vzorce", "Slovnik_vstupu", "Vstupy_a_KPI", "Konstanty_a_Ciselniky"];
const DEFAULT_WORKBOOK_NAME = "Seznam Vybavenosti_pracovni_katalog_pravidel.xlsx";
const PROJECT_SHEET_MARKER = "VYBAVENOST_PROJECT_V1";
const PROJECT_SHEET_CHUNK_SIZE = 30000;
const RESULTS_HEADING = "Hlavní výsledky návrhového výpočtu";
const MODEL_HEADERS = {
  id: "Kód / ID",
  name: "Typ zařízení a detailní vybavenost (Sub-sub skupina)",
  demand: "Kapacitní poptávka",
  capacityCount: "Kapacitní počet N_kapacita",
  spatialCount: "Plošná potřeba N_prostor",
  unit: "Měrná jednotka (MJ)",
  proposed: "Výsledný návrh N_navrh",
  areaPerPoint: "Průměrná plocha na bod",
  pavedArea: "Odhad zpevněné plochy areálu (m²)",
  unitPrice: "ZRN na jednotku (bez DPH)",
  cost: "Celkové náklady C_celk (Kč)",
  coverage: "Status prostorového pokrytí",
  energyReferenceArea: "Užitná plocha UP (m²; zdrojový cluster)",
  annualTotalEnergy: "Roční celková energie (kWh)",
  annualElectricity: "Roční elektřina (kWh)",
  peakElectricalLoad: "Odhad špičkového příkonu (kW)"
};
const RULE_HEADERS = {
  id: "ID pravidla",
  equipmentId: "ID vybavenosti",
  equipment: "Vybavenost",
  area: "Oblast",
  type: "Druh pravidla",
  condition: "Podmínka / rozsah",
  description: "Pravidlo nebo zjištění",
  formula: "Přesný vzorec / algoritmus",
  inputs: "Vstupy a jednotky",
  output: "Výstup a jednotka",
  rounding: "Zaokrouhlení",
  source: "Zdroj",
  location: "Konkrétní místo ve zdroji",
  url: "URL / soubor",
  verification: "Stav ověření",
  limitation: "Omezení / nutné rozhodnutí"
};
const INPUT_HEADERS = {
  id: "ID parametru",
  code: "Kód",
  name: "Název parametru",
  role: "Role v modelu",
  type: "Typ hodnoty",
  unit: "Jednotka",
  value: "Hodnota / výchozí stav v sešitu",
  validation: "Kontrola / povolené hodnoty",
  verification: "Stav ověření",
  derivation: "Vzorec / odvození",
  use: "Použití",
  editable: "Uživatelská editace",
  source: "Zdroj / poznámka"
};
const PRIMARY_INPUTS = [
  { row: 4, code: "O_celk", label: "Počet obyvatel řešeného území (přírůstek záměru)", kind: "number", min: 0, step: 1 },
  { row: null, code: "BJ_POCET_MANUAL", label: "Počet nových bytových jednotek (byty i rodinné domy)", kind: "number", min: 0, step: 1,
    defaultValue: "", unit: "bytových jednotek", note: "Ponechte prázdné pro odhad z obyvatel nebo z celkové HPP a průměrné plochy. Zahrnuje byty i rodinné domy." },
  { row: null, code: "PARKING_TOTAL_HPP", label: "Celková HPP navrženého bydlení (volitelné)", kind: "number", min: 0, step: 1,
    defaultValue: "", unit: "m² HPP", note: "Pokud není zadán počet jednotek, odhadne se jejich počet z HPP a vážené průměrné plochy podle typu zástavby." },
  { row: 5, code: "O_sidlo", label: "Současný počet obyvatel řešeného území", kind: "number", min: 0, step: 1,
    defaultValue: 1500 },
  { row: 6, code: "O_spad", label: "Počet obyvatel ve spádujících okolních sídlech (mimo řešené území, ne správní obvod)", kind: "number", min: 0, step: 1 },
  { row: 7, code: "S_uzemi", kind: "number", min: 0, step: 0.1 },
  { row: 8, code: "S_zastav", kind: "number", min: 0, step: 0.1 },
  { row: 11, code: "TYP_UZEMI", label: "Kategorie území dle intenzity využití", kind: "select",
    options: ["Typ A", "Typ B", "Typ C", "Typ D"],
    note: "Zvolte podle charakteru obce, rozvojového kontextu a oddělených částí sídla. Samotný počet obyvatel typ neurčuje; výchozí typ v sešitu je jen testovací klasifikace." },
  { row: 12, code: "KOMPAKT", kind: "select", options: ["1", "0"], labels: ["Ano", "Ne"] },
  { row: 13, code: "KOD_TVAR", label: "Geometrický tvar řešeného území (zjednodušený)", kind: "select",
    options: ["TVAR_1", "TVAR_2", "TVAR_3", "TVAR_4", "TVAR_5"],
    labels: ["Kruh", "Pravidelný pětiúhelník", "Čtverec", "Obdélník (poměr stran 1:2)", "Lineární pás (poměr stran 1:4)"] },
  { row: 14, code: "REZIM_STD", label: "Zvolená hladina navrhovaného standardu", kind: "select", options: ["ZAKON", "STANDARD"], labels: ["Zákonné minimum", "Doporučený standard"], unit: "" },
  { row: null, code: "PARKING_APARTMENT_SHARE_A", label: "Podíl bytových domů · typ A", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%" },
  { row: null, code: "PARKING_APARTMENT_SHARE_B", label: "Podíl bytových domů · typ B", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.8, unit: "%" },
  { row: null, code: "PARKING_APARTMENT_SHARE_C", label: "Podíl bytových domů · typ C", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.5, unit: "%" },
  { row: null, code: "PARKING_APARTMENT_SHARE_D", label: "Podíl bytových domů · typ D", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.2, unit: "%" },
  { row: null, code: "PARKING_LARGE_APARTMENT_SHARE", label: "Očekávaný podíl bytů nad 120 m²", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.1, unit: "%" },
  { row: null, code: "PARKING_SMALL_APARTMENT_AREA_M2", label: "Průměrná plocha bytu do 120 m²", kind: "number", min: 1, max: 120, step: 1, defaultValue: 65, unit: "m²" },
  { row: null, code: "PARKING_LARGE_APARTMENT_AREA_M2", label: "Průměrná plocha bytu nad 120 m²", kind: "number", min: 120, step: 1, defaultValue: 150, unit: "m²" },
  { row: null, code: "PARKING_FAMILY_HOME_AREA_M2", label: "Průměrná plocha rodinného domu", kind: "number", min: 1, step: 1, defaultValue: 140, unit: "m²" },
  { row: null, code: "PARKING_APARTMENT_RATE", label: "Modelový poměr stání na byt", kind: "number", min: 0, step: 0.1, defaultValue: 1.1, unit: "stání / byt" },
  { row: null, code: "PARKING_FAMILY_HOME_RATE", label: "Modelový poměr stání na rodinný dům", kind: "number", min: 0, step: 0.1, defaultValue: 2, unit: "stání / dům" },
  { row: null, code: "PARKING_1B_QUALIFIED", label: "Předpoklad splnění kategorie dostupného nájemního bydlení 1b", kind: "select",
    options: ["ANO", "NE"], labels: ["Ano — podíl městských bytů modelovat sazbou 1b", "Ne — všechno bydlení sazbou 1a"],
    defaultValue: "ANO", note: "Výchozí „Ano“ je scénářový předpoklad. Ověřte podmínky § 13 písm. q) zákona č. 283/2021 Sb.; samotný městský odkup za nákladovou cenu nemusí stačit." },
  { row: 16, code: "k_odkup_obec", kind: "percent", min: 0, max: 1, step: 0.01 },
  { row: 17, code: "m_marze_soc", kind: "percent", min: 0, max: 1, step: 0.01 }
];
const DERIVED_INPUTS = [
  { row: null, code: "O_uzemi_cil", label: "Počet obyvatel řešeného území po záměru", unit: "obyvatel", formula: "O_sidlo + O_celk" },
  { row: 19, code: "O_prah", label: "Počet obyvatel včetně spádujících okolních sídel", unit: "obyvatel", formula: "O_uzemi_cil + O_spad" },
  { row: null, code: "BJ_pocet", label: "Počet nových bytových jednotek", unit: "jednotek", formula: "Ruční počet; jinak celková HPP / vážená plocha jednotky; jinak O_celk / počet osob na byt" },
  { row: null, code: "HPP_byty", label: "Hrubá podlažní plocha nových bytů", unit: "m² HPP", formula: "BJ_pocet × HPP na byt" },
  { row: 20, code: "rho_hruba", label: "Hrubá obytná hustota", unit: "obyv./ha", formula: "O_celk / S_uzemi" },
  { row: 21, code: "rho_cista", label: "Čistá obytná hustota", unit: "obyv./ha", formula: "O_celk / S_zastav" },
  { row: 22, code: "k_tvar", label: "Tvarový koeficient", unit: "", formula: "Konstanty_a_Ciselniky: tvarový koeficient" },
  { row: 23, code: "k_pesi", label: "Uliční koeficient docházky", unit: "", formula: "Konstanty_a_Ciselniky: uliční koeficient" },
  { row: 24, code: "c_tvar", label: "Korekční činitel sítě bodů", unit: "", formula: "Konstanty_a_Ciselniky: korekce sítě bodů" },
  { row: 25, code: "R_max", label: "Maximální poloměr sítě", unit: "m", formula: "k_tvar × √(S_uzemi × 10 000)" },
  { row: 26, code: "L_max", label: "Max. pěší vzdálenost z těžiště", unit: "m", formula: "k_pesi × √(S_uzemi × 10 000)" }
];
const SECONDARY_INPUTS = [
  { row: 15, code: "HPP na byt", label: "Hrubá podlažní plocha jednoho bytu", kind: "number", min: 1, step: 1, max: 500, unit: "m² HPP / byt" },
  { row: 16, code: "Průměrná podlažnost", label: "Průměrný počet nadzemních podlaží", kind: "number", min: 1, step: 1, max: 100, unit: "podlaží" },
  { row: 17, code: "Koeficient zastavění", label: "Koeficient zastavění parcely", kind: "number", min: 0.01, step: 0.01, max: 1, unit: "zastavěná plocha / parcela" },
  { row: null, code: "Plocha pro odhad dostupnosti", label: "Plocha pro odhad rozmístění bydlení", kind: "select", options: ["S_zastav", "S_uzemi"], labels: ["Zastavitelné plochy bydlení a OV", "Celé řešené území"], defaultValue: "S_zastav", unit: "" },
  { row: null, code: "MAX_REZERVA", label: "Maximální % vyčerpání kapacity rezerv ve spádové oblasti", kind: "percent", min: 0, max: 1, step: 0.05, defaultValue: 0.5, unit: "%",
    note: "Výchozích 50 % je editovatelný scénář, nikoli právní ani metodická norma. Omezuje, kolik zadané kapacity okolí se započte do potřeby." }
];
const PROJECT_ESTIMATE_INPUTS = [
  {
    row: null, code: "K_CELKOVA_PLOCHA", label: "Koeficient celkové podlahové plochy / užitné plochy",
    kind: "number", min: 0.1, step: 0.05, defaultValue: 1,
    unit: "m² celkem / m² užitná",
    note: "Výchozí 1,00 ponechává užitnou plochu beze změny. Upravte podle požadované hrubé podlahové plochy; jde o editovatelný projektový předpoklad."
  },
  {
    row: null, code: "K_OBESTAVENY", label: "Koeficient obestavěného prostoru",
    kind: "number", min: 0.1, step: 0.1, defaultValue: 4.2,
    unit: "m³ / m² celkové podlahové plochy",
    note: "Výchozí 4,2 m³/m² je převzato ze scénáře polikliniky v pracovním zdravotnickém clusteru. Je to editovatelný odhad, ne univerzální norma; upravte jej i pro školy."
  },
  {
    row: null, code: "K_PLOCHA_AREALU", label: "Koeficient plochy areálu / celkové podlahové plochy",
    kind: "number", min: 0.1, step: 0.1, defaultValue: 750 / 280,
    unit: "m² areálu / m² celkem",
    note: "Výchozí poměr 750/280 vychází ze scénáře polikliniky v pracovním zdravotnickém clusteru. Upravte podle zadání; stejný koeficient lze použít i u škol."
  },
  {
    row: null, code: "K_ZASTAVENI", label: "Podíl zastavěné plochy / celkové podlahové plochy",
    kind: "number", min: 0.01, max: 1, step: 0.05, defaultValue: 0.5,
    unit: "m² půdorysu / m² celkem",
    note: "Výchozích 0,50 odpovídá dvoupodlažnímu scénáři v pracovním sešitu. Upravte podle zvolené podlažnosti."
  },
  {
    row: null, code: "K_ZPEVNENI", label: "Podíl zpevněné plochy ze zbytku areálu",
    kind: "percent", min: 0, max: 1, step: 0.05, defaultValue: 0.35,
    unit: "%; z plochy areálu mimo půdorys",
    note: "Výchozích 35 % je scénářový podíl v pracovním sešitu, nikoli zaměření. Upravte podle projektu."
  },
  {
    row: null, code: "K_CENA_SKOLY", label: "Koeficient investice školy vůči ceně modelu",
    kind: "number", min: 0.1, step: 0.05, defaultValue: 1,
    unit: "násobek",
    note: "Výchozí 1,00 ponechává cenu za m³ odvozenou z nákladů a objemu příslušné položky v Excelu. Upravte podle cenové úrovně či rozsahu projektu."
  },
  {
    row: null, code: "POLY_OBYV_NA_CENTRUM", label: "Obyvatel na jedno modelové poliklinické centrum",
    kind: "number", min: 1, step: 500, defaultValue: 10000,
    unit: "obyvatel / centrum",
    note: "Sešit používá hranici 10 000 obyvatel pro spuštění jednoho centra a označuje ji jako neověřenou. Aplikace z ní dělá poměrný scénář, ne zákonný práh."
  },
  {
    row: null, code: "POLY_UZITNA_PLOCHA", label: "Užitná plocha jednoho modelového poliklinického centra",
    kind: "number", min: 1, step: 10, defaultValue: 280,
    unit: "m² / centrum",
    note: "Výchozí hodnota pochází ze scénáře polikliniky v pracovním zdravotnickém clusteru. Jde o předpoklad pracovního modelu, nikoli právní minimum."
  },
  {
    row: null, code: "POLY_CENA_M3", label: "ZRN polikliniky na 1 m³ obestavěného prostoru",
    kind: "number", min: 1, step: 100, defaultValue: 13645,
    unit: "Kč / m³",
    note: "Sazba RTS se načítá podle JKSO 801.1 a zvolené konstrukční charakteristiky. Pokud RTS pro kombinaci cenu neuvádí, zadejte ověřenou sazbu ručně."
  },
  {
    row: null, code: "POLY_PRIRAZKA", label: "Přirážka VRN a rezervy k ZRN",
    kind: "percent", min: 0, max: 2, step: 0.01, defaultValue: 0.23,
    unit: "%; z ZRN",
    note: "Výchozích 23 % tvoří 8 % VRN + 15 % rezerva podle pracovního číselníku. Jde o editovatelný modelový předpoklad."
  }
];
const HOUSING_CONTRIBUTION_INPUTS = [
  {
    row: null, code: "BYTY_PLOCHA_NA_BYT", label: "Užitná plocha jednoho bytu",
    kind: "number", min: 1, step: 1, defaultValue: 52,
    unit: "m² / byt",
    note: "Pracovní předpoklad bytového scénáře; výchozí koeficient je odvozen z modelu sociální vybavenosti a bydlení."
  },
  {
    row: null, code: "BYTY_OBESTAVENI_NA_M2", label: "Obestavěný prostor na 1 m² bytové plochy",
    kind: "number", min: 0.1, step: 0.1, defaultValue: 3.9,
    unit: "m³ / m²",
    note: "Pracovní předpoklad bytového scénáře; výchozí koeficient je odvozen z modelu sociální vybavenosti a bydlení."
  },
  {
    row: null, code: "BYTY_CENA_M3", label: "Modelová hodnota na 1 m³ obestavěného prostoru",
    kind: "number", min: 1, step: 100, defaultValue: 12400,
    unit: "Kč / m³",
    note: "Výchozí sazba 12 400 Kč/m³ je uvedena u bytového scénáře; nejde o ověřenou tržní prodejní cenu."
  }
];
const ITEM_ESTIMATE_INPUTS = [
  { code: "ITEM_QUANTITY", label: "Počet jednotek v propočtu", kind: "number", min: 0, step: 1, unit: "jednotek", color: "area" },
  { code: "ITEM_USABLE_AREA_PER_UNIT", label: "Užitná plocha na jednotku", kind: "number", min: 0, step: 0.1, unit: "m² / jednotku", color: "area" },
  { code: "ITEM_TOTAL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, unit: "m² celkem / m² užitná", color: "floor" },
  { code: "ITEM_VOLUME_PER_AREA", label: "Koeficient obestavěného prostoru", kind: "number", min: 0, step: 0.1, unit: "m³ / m² celkem", color: "volume" },
  { code: "ITEM_SITE_AREA_PER_UNIT", label: "Plocha pozemku na jednotku", kind: "number", min: 0, step: 0.1, unit: "m² / jednotku", color: "site" },
  { code: "ITEM_PAVED_AREA_PER_UNIT", label: "Zpevněná plocha na jednotku", kind: "number", min: 0, step: 0.1, unit: "m² / jednotku", color: "paved" },
  { code: "ITEM_RTS_RATE", label: "Cena dle měrné jednotky", kind: "number", min: 0, step: 100, unit: "Kč / měrná jednotka", color: "rts" },
  { code: "ITEM_RTS_RATE_BASIS", label: "Základ ceny", kind: "select", options: ["volume", "unit"], labels: ["Kč / m³ obestavěného prostoru", "Kč / měrnou jednotku"], unit: "", color: "rts" },
  { code: "ITEM_RTS_FIXED_COST", label: "Další fixní náklady dle modelu", kind: "number", min: 0, step: 10000, unit: "Kč / projekt", color: "fixed" },
  { code: "ITEM_COST_SURCHARGE", label: "VRN a rezerva", kind: "percent", min: 0, max: 2, step: 0.01, unit: "%; z nákladů ZRN", color: "surcharge" }
];
const ALTERNATIVE_ESTIMATE_INPUTS = [
  { code: "ALT_RAINFALL_DEPTH_MM", label: "Modelová srážka", kind: "number", min: 0, step: 1, unit: "mm / srážková událost", color: "area" },
  { code: "ALT_RUNOFF_COEFFICIENT", label: "Součinitel odtoku", kind: "number", min: 0, max: 1, step: 0.05, unit: "bez jednotky", color: "paved" },
  { code: "ALT_RAINWATER_RATE", label: "Odhad ceny retenční nádrže včetně stavebních prací", kind: "number", min: 0, step: 100, unit: "Kč / m³ retenčního objemu", color: "rts" },
  { code: "ALT_WATER_LITERS_PER_PERSON_DAY", label: "Spotřeba vody na obyvatele", kind: "number", min: 0, step: 1, unit: "l / osoba / den", color: "area" },
  { code: "ALT_PEAK_DAY_FACTOR", label: "Součinitel špičkového dne", kind: "number", min: 1, step: 0.1, unit: "násobek průměrného dne", color: "volume" },
  { code: "ALT_WASTEWATER_RETURN", label: "Podíl vody odváděný jako splašky", kind: "percent", min: 0, max: 1, step: 0.01, unit: "% z dodané vody", color: "paved" },
  { code: "ALT_NETWORK_LENGTH_M", label: "Délka nové trasy potrubí nebo kabelu", kind: "number", min: 0, step: 50, unit: "m", color: "site" },
  { code: "ALT_WATER_PIPE_RATE", label: "Odhad ceny vodovodního potrubí včetně pokládky", kind: "number", min: 0, step: 100, unit: "Kč / m", color: "rts" },
  { code: "ALT_WATER_METER_CHAMBER_COUNT", label: "Počet vodoměrných a armaturních šachet", kind: "number", min: 0, step: 1, unit: "ks", color: "fixed" },
  { code: "ALT_WATER_METER_CHAMBER_RATE", label: "Odhad ceny jedné vodoměrné a armaturní šachty", kind: "number", min: 0, step: 1000, unit: "Kč / ks", color: "rts" },
  { code: "ALT_SEWER_PIPE_RATE", label: "Odhad ceny splaškového potrubí včetně pokládky", kind: "number", min: 0, step: 100, unit: "Kč / m", color: "rts" },
  { code: "ALT_SEWER_MANHOLE_SPACING_M", label: "Modelová rozteč revizních šachet", kind: "number", min: 1, step: 5, unit: "m", color: "site" },
  { code: "ALT_SEWER_MANHOLE_RATE", label: "Odhad ceny jedné revizní šachty", kind: "number", min: 0, step: 1000, unit: "Kč / ks", color: "rts" },
  { code: "ALT_ENERGY_DENSITY_W_M2", label: "Modelová elektrická zátěž budov", kind: "number", min: 0, step: 5, unit: "W / m² celkové plochy", color: "volume" },
  { code: "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO", label: "Poměr špičkového a průměrného elektrického příkonu", kind: "number", min: 1, step: 0.1, unit: "špička / průměr", color: "volume" },
  { code: "BOILER_HEAT_DENSITY_KWH_M2_Y", label: "Náhradní roční potřeba tepla pro objekty bez energetického údaje", kind: "number", min: 0, step: 5, unit: "kWh / m² / rok", color: "volume" },
  { code: "BOILER_HEATING_SHARE", label: "Podíl neelektrické energie uvažovaný pro ústřední vytápění", kind: "percent", min: 0, max: 1, step: 0.05, unit: "% neelektrické energie", color: "area" },
  { code: "BOILER_FULL_LOAD_HOURS", label: "Ekvivalentní roční hodiny plného zatížení kotelny", kind: "number", min: 1, step: 100, unit: "h / rok", color: "volume" },
  { code: "BOILER_DESIGN_RESERVE", label: "Součinitel návrhové rezervy výkonu", kind: "number", min: 1, step: 0.05, unit: "násobek", color: "volume" },
  { code: "BOILER_UNIT_CAPACITY_KW", label: "Kapacita jednoho kotlového modulu", kind: "number", min: 1, step: 50, unit: "kW tepelného výkonu / modul", color: "volume" },
  { code: "BOILER_UNIT_COST", label: "Modelový odhad ceny jednoho kotlového modulu", kind: "number", min: 0, step: 10000, unit: "Kč / modul; odhad 10 000 Kč/kW", color: "rts" },
  { code: "ALT_POWER_CABLE_RATE", label: "Odhad ceny elektrického kabelu včetně pokládky", kind: "number", min: 0, step: 100, unit: "Kč / m", color: "rts" },
  { code: "ALT_TRANSFORMER_CAPACITY_KW", label: "Modelová kapacita jedné trafostanice", kind: "number", min: 1, step: 100, unit: "kW / stanici", color: "volume" },
  { code: "ALT_TRANSFORMER_RATE", label: "Odhad ceny jedné trafostanice včetně osazení", kind: "number", min: 0, step: 10000, unit: "Kč / stanici", color: "rts" },
  { code: "ALT_COST_SURCHARGE", label: "VRN a rezerva", kind: "percent", min: 0, max: 2, step: 0.01, unit: "%; z nákladů ZRN", color: "surcharge" }
];
const ITEM_ESTIMATE_SCENARIOS = {
  VYB_05B: { usableArea: 1800, floorFactor: 1.1, volumeFactor: 4.2, siteArea: 6000, pavedArea: 3000 },
  VYB_09B_1: { usableArea: 90, floorFactor: 1.1, volumeFactor: 3.3, siteArea: 250, pavedArea: 100 },
  VYB_09B_2: { usableArea: 90, floorFactor: 1.1, volumeFactor: 3.3, siteArea: 250, pavedArea: 100 },
  VYB_10B_1: { usableArea: 250, floorFactor: 1.1, volumeFactor: 3.3, siteArea: 900, pavedArea: 450 },
  VYB_11: { usableArea: 300, floorFactor: 1.1, volumeFactor: 3.3, siteArea: 900, pavedArea: 450 },
  VYB_13A: { usableArea: 250, floorFactor: 1.05, volumeFactor: 3, siteArea: 900, pavedArea: 450 },
  VYB_13B: { usableArea: 1500, floorFactor: 1.05, volumeFactor: 3, siteArea: 5000, pavedArea: 2500 }
};
const HEATING_BOILER_ID = "VYB_20B_4";
const RESIDENT_PARKING_ID = "VYB_15A_1";
const RESIDENT_PARKING_LAW_ID = "VYB_15A_1_LAW";
const RESIDENT_PARKING_RESERVE_ID = "VYB_15A_1_RESERVE";
const RESIDENT_PARKING_IDS = new Set([
  RESIDENT_PARKING_ID, RESIDENT_PARKING_LAW_ID, RESIDENT_PARKING_RESERVE_ID
]);
const TECHNICAL_ALTERNATIVE_IDS = new Set(["VYB_20A", "VYB_20B_1", "VYB_20B_2", "VYB_20B_3", HEATING_BOILER_ID]);
const INFRASTRUCTURE_ITEM_IDS = TECHNICAL_ALTERNATIVE_IDS;
const TECHNICAL_INFRASTRUCTURE_CATEGORY = "TECHNICKÁ INFRASTRUKTURA A VODNÍ HOSPODÁŘSTVÍ";
const RECOMMENDED_SCENARIO_ITEM_IDS = new Set([
  "VYB_11", "VYB_01_J", "VYB_02C_1", "VYB_02C_2", "VYB_09A_4",
  "VYB_09D", "VYB_16A", "VYB_17A"
]);
const REALIZATION_NETWORK_IDS = new Set(["VYB_20B_1", "VYB_20B_2", "VYB_20B_3"]);
const REALIZATION_PARKING_IDS = new Set([
  "VYB_15A_1", RESIDENT_PARKING_LAW_ID, RESIDENT_PARKING_RESERVE_ID,
  "VYB_15A_2", "VYB_15A_3", "VYB_15B_1", "VYB_15B_2"
]);
const REALIZATION_NON_BUILDING_IDS = new Set(["VYB_14", "VYB_07B"]);
const REVIEW_ALTERNATIVE_IDS = new Set([
  "VYB_11", "VYB_13A", "VYB_13B",
  ...TECHNICAL_ALTERNATIVE_IDS
]);
const REVIEW_ALTERNATIVE_LABELS = {
  VYB_11: {
    name: "Úřad a kontaktní místo samosprávy",
    context: "Prověřte kapacitu a dostupnost stávající veřejné správy. Nejde o požadavek vybudovat nový úřad."
  },
  VYB_13A: {
    name: "Alternativa každodenního nákupu – menší prodejna",
    context: "Odděluje průběžný místní nákup od většího týdenního nákupu v supermarketu."
  },
  VYB_13B: {
    name: "Alternativa týdenního nákupu – supermarket",
    context: "Jde o větší nákupní formát; nenahrazuje menší prodejnu ani se s ní nesčítá jako stejná provozovna."
  },
  VYB_20A: {
    name: "Retenční nádrž na dešťovou vodu – cena za m³ objemu",
    context: "Odhad ceny nádrže a stavebních prací vychází z vypočteného objemu zachycené srážkové vody."
  },
  VYB_20B_1: {
    name: "Vodovodní potrubí – cena za metr a vodoměrná šachta",
    context: "Odhaduje dodávku a pokládku potrubí podle délky trasy, průtok podle počtu obyvatel a cenu vodoměrných či armaturních šachet."
  },
  VYB_20B_2: {
    name: "Splašková kanalizace – cena potrubí za metr a revizních šachet",
    context: "Odhaduje gravitační potrubí podle délky trasy, odtok podle spotřeby vody a revizní šachty v modelové rozteči."
  },
  VYB_20B_3: {
    name: "Elektrické kabelové vedení – cena za metr a trafostanice",
    context: "Odhaduje potřebný příkon budov, pokládku kabelového vedení podle délky trasy a počet trafostanic podle jejich modelové kapacity."
  }
};
const TECHNICAL_CATALOG_NAMES = {
  VYB_20A: REVIEW_ALTERNATIVE_LABELS.VYB_20A.name,
  VYB_20B_1: REVIEW_ALTERNATIVE_LABELS.VYB_20B_1.name,
  VYB_20B_2: REVIEW_ALTERNATIVE_LABELS.VYB_20B_2.name,
  VYB_20B_3: REVIEW_ALTERNATIVE_LABELS.VYB_20B_3.name
};
const PROJECT_LOCATION_INPUT = {
  code: "OBEC",
  label: "Obec / městská část pro ověření",
  kind: "text",
  defaultValue: "Brno-Bystrc"
};
const TERRITORY_TYPE_HELP_URL = "TB050MMR01-Standardy-dostupnosti-verejne-infrastruktury-aktualizace-2020-03.pdf";
const COVERAGE_OVERLAP_INPUT = {
  code: "PREKRYTI",
  label: "Bezpečnostní překrytí dosahových kružnic",
  kind: "percent",
  min: 0,
  max: 0.75,
  step: 0.05,
  defaultValue: 0.25,
  unit: "%; zkrácení kroku sítě"
};
const HALL_COVERAGE_INPUT = {
  code: "ITEM_HALL_COVERAGE",
  label: "Veřejná krytá hala – způsob pokrytí",
  kind: "select",
  row: 23,
  options: [
    "Neověřeno", "Kapacita ZŠ/okolí pokrývá potřebu", "Není samostatná hala požadována", "Samostatná nová hala požadována"
  ],
  unit: "ověření kapacity"
};
const JKSO_CONSTRUCTION_INPUT = {
  code: "ITEM_CONSTRUCTION_SYSTEM",
  label: "Konstrukční systém objektu dle RTS / JKSO",
  kind: "select",
  options: ["1", "2", "3", "4", "5", "6", "7", "8", "9"],
  labels: [
    "1 — zděná z cihel, tvárnic nebo bloků",
    "2 — monolitická betonová tyčová",
    "3 — monolitická betonová plošná",
    "4 — montovaná z betonových tyčových dílců",
    "5 — montovaná z betonových plošných dílců",
    "6 — montovaná z prostorových buněk",
    "7 — kovová",
    "8 — dřevěná nebo na bázi dřeva",
    "9 — z jiných materiálů"
  ],
  unit: "charakteristika konstrukce"
};
const RTS_801_CONSTRUCTION_RATES = {
  "801.1": { "1": 13645, "2": 14020, "3": 13850, "4": 13060 },
  "801.2": { "1": 14800, "2": 14070, "4": 18810, "5": 13100, "6": 13230 },
  "801.3": { "1": 12385, "2": 9170, "4": 11650, "5": 11650, "6": 10610, "8": 18850 },
  "801.4": { "1": 14760, "2": 9200, "3": 16350, "4": 20500, "5": 14660, "8": 13100 },
  "801.5": { "1": 13250, "2": 11520, "3": 16830, "4": 9440, "8": 15200 },
  "801.6": { "1": 12255, "2": 11010, "3": 12830, "5": 12290, "6": 12570, "7": 7370, "8": 15010, "9": 14710 },
  "801.7": { "1": 12075, "2": 11270, "3": 11900, "4": 12830, "5": 16690, "6": 9570, "9": 10190 },
  "801.8": { "1": 11625, "2": 11370, "4": 13050, "8": 10460 },
  "801.9": { "1": 10520, "2": 13760, "5": 11510, "6": 9440, "7": 7360 }
};
const RTS_801_GROUP_BY_ITEM_ID = {
  VYB_01_J: "801.3", VYB_01: "801.3", VYB_02A: "801.3", VYB_02B: "801.3",
  VYB_02C_1: "801.3", VYB_02C_2: "801.3",
  VYB_05B: "801.5",
  VYB_06A: "801.9", VYB_06B: "801.9", VYB_07A_1: "801.9", VYB_07A_2: "801.9",
  VYB_09B_1: "801.1", VYB_09B_2: "801.1", VYB_09A_4: "801.1", VYB_16: "801.1",
  VYB_17: "801.1", VYB_18: "801.1", VYB_19: "801.1", VYB_20: "801.1",
  VYB_09D: "801.6", VYB_16A: "801.6", VYB_17A: "801.6",
  VYB_09D: "801.6", VYB_16A: "801.6", VYB_17A: "801.6",
  VYB_10A: "801.4", VYB_10B_1: "801.4", VYB_10B_2: "801.4",
  VYB_11: "801.6",
  VYB_13A: "801.8", VYB_13B: "801.8"
};
const RTS_801_SOURCE = "RTS, Cenové ukazatele ve stavebnictví pro rok 2026, s. 1, skupina JKSO 801; ceny v Kč/m³ obestavěného prostoru (m³ OP).";
const PROJECT_INPUT_GROUPS = [
  {
    label: "Parametry projektu",
    fields: [
      { code: "REZIM_STD" },
      { field: PROJECT_LOCATION_INPUT },
      { code: "k_odkup_obec" },
      { code: "m_marze_soc" },
      { code: "MAX_REZERVA", secondary: true }
    ]
  },
  {
    label: "Plochy a tvar území",
    fields: [
      { code: "S_uzemi" }, { code: "S_zastav" }, { code: "TYP_UZEMI" }, { code: "KOMPAKT" },
      { code: "KOD_TVAR" }, { field: COVERAGE_OVERLAP_INPUT, secondary: true }
    ]
  },
  {
    label: "Počty obyvatel",
    fields: [{ code: "O_sidlo" }, { code: "O_celk" }, { code: "PARKING_TOTAL_HPP" }, { code: "O_spad" }]
  },
  {
    label: "Parkování podle skladby bydlení",
    fields: [
      { code: "PARKING_APARTMENT_SHARE_A" }, { code: "PARKING_APARTMENT_SHARE_B" },
      { code: "PARKING_APARTMENT_SHARE_C" }, { code: "PARKING_APARTMENT_SHARE_D" },
      { code: "PARKING_APARTMENT_RATE" }, { code: "PARKING_FAMILY_HOME_RATE" },
      { code: "PARKING_1B_QUALIFIED" }
    ]
  },
  {
    label: "Počet bytů, podlažnost, HPP a plošné předpoklady",
    fields: [
      { code: "BJ_POCET_MANUAL" },
      { code: "PARKING_LARGE_APARTMENT_SHARE" },
      { code: "PARKING_SMALL_APARTMENT_AREA_M2" },
      { code: "PARKING_LARGE_APARTMENT_AREA_M2" },
      { code: "PARKING_FAMILY_HOME_AREA_M2" },
      { code: "Průměrná podlažnost", secondary: true },
      { code: "HPP na byt", secondary: true },
      { code: "Koeficient zastavění", secondary: true }
    ]
  }
];
const GENERAL_OBLIGATION_RULE_IDS = new Set(["ZEL-LEG-002", "SAFE-LEG-001", "TECH-LEGAL-002"]);
const REVIEW_ONLY_ITEM_IDS = new Set(["VYB_16B"]);
const RULE_MODEL_LINKS = {
  "ZEL-MET-001": ["VYB_03A"],
  "ZEL-MET-002": ["VYB_03B"],
  "ZEL-MET-003": ["VYB_03C"],
  "ZEL-MET-004": ["VYB_04B"],
  "ZEL-MET-005": ["VYB_04", "VYB_05B"],
  "ZEL-GEO-001": ["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04B"],
  "ZEL-MODEL-001": ["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04B"],
  "ZEL-ASSUM-001": ["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04", "VYB_04B", "VYB_05B"],
  "ZEL-RTS-001": ["VYB_05B"],
  "ZEL-RTS-002": ["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04B"],
  "ZEL-CAP-001": ["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04", "VYB_04B", "VYB_05B"],
  "ZEL-DATA-001": ["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04A", "VYB_04", "VYB_04B", "VYB_05B"],
  "ZDR-LEG-001": ["VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20", "VYB_09A_4"],
  "ZDR-MET-001": ["VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20"],
  "ZDR-SPEC-001": ["VYB_09B_1", "VYB_09B_2"],
  "ZDR-BLDG-001": ["VYB_17", "VYB_18", "VYB_19", "VYB_20"],
  "ZDR-CAP-001": ["VYB_17", "VYB_18", "VYB_19", "VYB_20", "VYB_09A_4"],
  "ZDR-CAP-002": ["VYB_17", "VYB_18", "VYB_19", "VYB_20", "VYB_09A_4"],
  "ZDR-RTS-001": ["VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20"],
  "SOC-LEG-001": ["VYB_06A", "VYB_06B", "VYB_06C", "VYB_07A_1", "VYB_07A_2"],
  "SOC-MET-001": ["VYB_06A"],
  "SOC-MET-002": ["VYB_06A", "VYB_06B", "VYB_06C", "VYB_07A_1", "VYB_07A_2"],
  "SOC-MET-003": ["VYB_06A"],
  "SOC-CAP-001": ["VYB_06A", "VYB_06B", "VYB_06C", "VYB_07A_1", "VYB_07A_2"],
  "SOC-DATA-001": ["VYB_06A", "VYB_06B", "VYB_06C", "VYB_07A_1", "VYB_07A_2"],
  "CULT-LAW-001": ["VYB_10A", "VYB_10B_1", "VYB_10B_2", "VYB_11", "VYB_12", "VYB_13A", "VYB_13B"],
  "CULT-MET-001": ["VYB_10B_2"],
  "CULT-MET-002": ["VYB_10A"],
  "CULT-MET-003": ["VYB_10B_1"],
  "CULT-MET-004": ["VYB_12"],
  "CULT-COMM-001": ["VYB_13A", "VYB_13B"],
  "MOB-MET-001": ["VYB_14"],
  "MOB-MET-002": ["VYB_15A_1", "VYB_15A_2", "VYB_15A_3"],
  "MOB-MET-003": ["VYB_15B_1", "VYB_15B_2"],
  "MOB-MET-004": ["VYB_15B_2"],
  "SAFE-MET-001": ["VYB_16B"],
  "SAFE-LEG-001": ["VYB_17B"],
  "TECH-LEGAL-002": ["VYB_18A", "VYB_18B"],
  "TECH-MET-001": ["VYB_18A"],
  "TECH-MET-002": ["VYB_18B"],
  "TECH-MET-003": ["VYB_19"],
  "TECH-HYDRO-001": ["VYB_20A"],
  "TECH-COST-001": ["VYB_20B_1", "VYB_20B_2", "VYB_20B_3"],
  "TECH-MET-005": ["VYB_18B", "VYB_19"],
};
const CAPACITY_QUANTITY_IDS = new Set([
  "VYB_06A", "VYB_06B", "VYB_06C", "VYB_07A_1", "VYB_07A_2",
  "VYB_16", "VYB_17", "VYB_18", "VYB_19",
  "VYB_15A_1", "VYB_15A_2", "VYB_15A_3", "VYB_15B_1", "VYB_15B_2"
]);
const CAPACITY_DEMAND_UNITS = {
  VYB_01: "děti",
  VYB_02A: "žáci 1. stupně",
  VYB_02B: "žáci 2. stupně",
  VYB_04A: "m²"
};
const SCHOOL_CAPACITY_ITEM_IDS = new Set(["VYB_01", "VYB_02A", "VYB_02B"]);
const DEMAND_LABELS = {
  VYB_01: "dětí",
  VYB_02A: "žáků 1. stupně",
  VYB_02B: "žáků 2. stupně"
};
const DEMAND_SEARCH_LABELS = {
  VYB_01: "děti",
  VYB_02A: "žáci 1. stupně",
  VYB_02B: "žáci 2. stupně"
};
const SPATIAL_ESTIMATE_CONFIG = {
  VYB_03A: { ruleId: "ZEL-MET-001", distanceMeters: 200 },
  VYB_03B: { ruleId: "ZEL-MET-002", distanceMeters: 500 },
  VYB_03C: { ruleId: "ZEL-MET-003", distanceMeters: "by-settlement-type" },
  VYB_04B: { ruleId: "ZEL-MET-004", distanceMeters: "by-settlement-type" }
};
const GEOMETRY_EXTENTS = {
  TVAR_1: { aspect: 1, areaFactor: 4 / Math.PI },
  TVAR_2: { aspect: 1.0514, areaFactor: 1.447 },
  TVAR_3: { aspect: 1, areaFactor: 1 },
  TVAR_4: { aspect: 2, areaFactor: 1 },
  TVAR_5: { aspect: 4, areaFactor: 1 }
};
const ITEM_DISPLAY_NAMES = {
  VYB_03A: "Hřiště pro předškolní děti",
  VYB_03B: "Hřiště pro mladší školní děti",
  VYB_03C: "Sportoviště pro mládež a dospělé",
  VYB_04A: "Veřejná prostranství – zákonné minimum",
  VYB_04B: "Veřejná parková zeleň",
  VYB_05B: "Krytá veřejná sportovní hala",
  VYB_10B_1: "Víceúčelový kulturní sál a kulturní dům (společenský sál)",
  VYB_11: "Úřad a kontaktní místo samosprávy",
  VYB_14: "Zastávky veřejné dopravy – pěší dostupnost",
  VYB_16: "Lékárna",
  VYB_15A_1: "Rezidentské parkování",
  VYB_15A_2: "Parkovací stání pro sociální služby",
  VYB_15A_3: "Návštěvnická parkovací stání",
  VYB_15B_1: "P+R / přestupní parkoviště",
  VYB_15B_2: "Cyklistická stání a B+R"
};
const TARGETED_PLAN_INPUTS = {
  VYB_01_J: [
    { code: "MODEL_DEMOGRAPHIC_SHARE", label: "Podíl dětí ve věku 0–3 roky", kind: "percent", min: 0, max: 1, step: 0.001, defaultValue: 0.042, unit: "%", note: "Výchozí scénář; nahraďte věkovou strukturou ČSÚ pro spád." },
    { code: "MODEL_ATTENDANCE_SHARE", label: "Podíl kohorty využívající péči", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.3, unit: "%", note: "Plánovací scénář, nikoli normativ." },
    { code: "MODEL_GROUP_CAPACITY", label: "Kapacita jedné dětské skupiny", kind: "number", min: 1, step: 1, defaultValue: 24, unit: "míst" },
    { code: "MODEL_MIN_GROUP_CAPACITY", label: "Minimální návrhová kapacita skupiny", kind: "number", min: 1, step: 1, defaultValue: 12, unit: "míst" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na jedno místo", kind: "number", min: 0, step: 0.1, defaultValue: 12.5, unit: "m²/místo" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jedno místo", kind: "number", min: 0, step: 0.1, defaultValue: 35, unit: "m²/místo" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.3 (výukové stavby)", kind: "number", min: 0, step: 1, defaultValue: 12385, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "EARLY_CHILDCARE_EXISTING_CAPACITY", label: "Ověřená volná kapacita okolních zařízení", kind: "number", min: 0, step: 1, unit: "míst" },
    { code: "EARLY_CHILDCARE_ACCESS_DISTANCE", label: "Skutečná pěší vzdálenost k zařízení s volnou kapacitou", kind: "number", min: 0, step: 1, unit: "m" },
    { code: "EARLY_CHILDCARE_DATA_SOURCE", label: "Zdroj demografie a kapacit", kind: "text", unit: "" }
  ],
  VYB_09A_4: [
    { code: "MODEL_POPULATION_PER_FTE", label: "Pacientek na 1 jednotku souč. zátěže (doporučeno 14 000)", kind: "number", min: 1, step: 100, defaultValue: 14000, unit: "pacientek/jednotku souč. zátěže", note: "Historický indikátor TAČR/MMR z podkladů ÚZIS a VZP, nikoli norma ČSÚ. Upravte podle místní věkové struktury a poskytovatelů." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jednoho ordinačního modulu", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "souč. zátěže" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na jednotku souč. zátěže", kind: "number", min: 0, step: 1, defaultValue: 120, unit: "m²/jednotku souč. zátěže" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jednotku souč. zátěže", kind: "number", min: 0, step: 1, defaultValue: 360, unit: "m²/jednotku souč. zátěže" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "GYNECOLOGY_AVAILABLE_FTE", label: "Ověřená volná souč. zátěž gynekologie", kind: "number", min: 0, step: 0.1, unit: "souč. zátěže" },
    { code: "GYNECOLOGY_WALK_DISTANCE", label: "Skutečná pěší vzdálenost k ordinaci", kind: "number", min: 0, step: 1, unit: "m" },
    { code: "GYNECOLOGY_DRIVE_MINUTES", label: "Ověřený dojezd (typ D)", kind: "number", min: 0, step: 0.1, unit: "min" },
    { code: "GYNECOLOGY_ACCESS_SOURCE", label: "Zdroj / datum ověření kapacity a dostupnosti", kind: "text", unit: "" }
  ],
  VYB_02C_1: [
    { code: "MODEL_DEMOGRAPHIC_SHARE", label: "Podíl cílové středoškolské kohorty", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.05, unit: "%", note: "Scénářový odhad; přednost má věková struktura a dojížďka v celém spádu." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné střední školy", kind: "number", min: 1, step: 1, defaultValue: 600, unit: "žákovských míst" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na žákovské místo", kind: "number", min: 0, step: 0.1, defaultValue: 10, unit: "m²/místo" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na žákovské místo", kind: "number", min: 0, step: 0.1, defaultValue: 25, unit: "m²/místo" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.3 (výukové stavby)", kind: "number", min: 0, step: 1, defaultValue: 12385, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "SS_EXISTING_FREE_SEATS", label: "Ověřená volná kapacita stávajících SŠ", kind: "number", min: 0, step: 1, unit: "míst" },
    { code: "SS_VHD_MINUTES", label: "Ověřený čas VHD k dostupné SŠ (typ A/B)", kind: "number", min: 0, step: 0.1, unit: "min" },
    { code: "SS_CAPACITY_SOURCE", label: "Zdroj demografie, kapacit a dostupnosti", kind: "text", unit: "" }
  ],
  VYB_02C_2: [
    { code: "MODEL_DEMOGRAPHIC_SHARE", label: "Podíl obyvatel využívajících ZUŠ", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.04, unit: "%", note: "Scénářový odhad; nahraďte místními daty o účasti a spádovosti." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné ZUŠ", kind: "number", min: 1, step: 1, defaultValue: 300, unit: "žákovských míst" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na žákovské místo", kind: "number", min: 0, step: 0.1, defaultValue: 8, unit: "m²/místo" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na žákovské místo", kind: "number", min: 0, step: 0.1, defaultValue: 12, unit: "m²/místo" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.4 (kultura a osvěta; proxy)", kind: "number", min: 0, step: 1, defaultValue: 14760, unit: "Kč/m³", note: "Cenový proxy pro vzdělávání v ZUŠ; upravte podle konkrétního stavebního programu." },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "ZUS_EXISTING_FREE_SEATS", label: "Ověřená volná kapacita stávajících ZUŠ", kind: "number", min: 0, step: 1, unit: "míst" },
    { code: "ZUS_WALK_DISTANCE", label: "Skutečná pěší vzdálenost k dostupné ZUŠ", kind: "number", min: 0, step: 1, unit: "m" },
    { code: "ZUS_VHD_MINUTES", label: "Ověřený čas VHD v mikroregionu", kind: "number", min: 0, step: 0.1, unit: "min" },
    { code: "ZUS_CAPACITY_SOURCE", label: "Zdroj účasti, kapacit a dostupnosti", kind: "text", unit: "" }
  ],
  VYB_09D: [
    { code: "MODEL_POPULATION_PER_VEHICLE", label: "Obyvatel na 1 výjezdový vůz (scénář)", kind: "number", min: 1, step: 1000, defaultValue: 50000, unit: "obyvatel/vůz" },
    { code: "MODEL_VEHICLES_PER_BASE", label: "Výjezdové vozy v jedné základně", kind: "number", min: 1, step: 1, defaultValue: 2, unit: "vozů" },
    { code: "MODEL_BASE_AREA", label: "Základní plocha jedné základny", kind: "number", min: 0, step: 1, defaultValue: 450, unit: "m²/základnu" },
    { code: "MODEL_AREA_PER_UNIT", label: "Plocha na jeden výjezdový vůz", kind: "number", min: 0, step: 1, defaultValue: 90, unit: "m²/vůz" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jednu základnu", kind: "number", min: 0, step: 1, defaultValue: 1800, unit: "m²/základnu" },
    { code: "MODEL_PAVED_PER_BASE", label: "Zpevněná plocha základny", kind: "number", min: 0, step: 1, defaultValue: 650, unit: "m²/základnu" },
    { code: "MODEL_PAVED_PER_UNIT", label: "Zpevněná plocha na výjezdový vůz", kind: "number", min: 0, step: 1, defaultValue: 100, unit: "m²/vůz" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 4, unit: "m³/m²" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.6 (provozní stavby; proxy)", kind: "number", min: 0, step: 1, defaultValue: 12255, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "ZZS_COVERED_POPULATION", label: "Obyvatelé projektu a spádu dosažitelní ze stávající základny do 20 min", kind: "number", min: 0, step: 1, unit: "obyvatel" },
    { code: "ZZS_MAX_RESPONSE_MINUTES", label: "Ověřená nejdelší dojezdová doba v projektovém spádu", kind: "number", min: 0, step: 0.1, unit: "min" },
    { code: "ZZS_COVERAGE_SOURCE", label: "Zdroj / datum ověření krajského plánu nebo GIS analýzy", kind: "text", unit: "" }
  ],
  VYB_16A: [
    { code: "MODEL_POPULATION_PER_VEHICLE", label: "Obyvatel na 1 požární vozidlo (scénář)", kind: "number", min: 1, step: 1000, defaultValue: 10000, unit: "obyvatel/vozidlo" },
    { code: "MODEL_VEHICLES_PER_BASE", label: "Požární vozidla v jedné stanici", kind: "number", min: 1, step: 1, defaultValue: 4, unit: "vozidel" },
    { code: "MODEL_BASE_AREA", label: "Základní plocha jedné stanice", kind: "number", min: 0, step: 1, defaultValue: 650, unit: "m²/stanici" },
    { code: "MODEL_AREA_PER_UNIT", label: "Plocha na jedno požární vozidlo", kind: "number", min: 0, step: 1, defaultValue: 75, unit: "m²/vozidlo" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jednu stanici", kind: "number", min: 0, step: 1, defaultValue: 4000, unit: "m²/stanici" },
    { code: "MODEL_PAVED_PER_BASE", label: "Zpevněná plocha stanice", kind: "number", min: 0, step: 1, defaultValue: 1200, unit: "m²/stanici" },
    { code: "MODEL_PAVED_PER_UNIT", label: "Zpevněná plocha na požární vozidlo", kind: "number", min: 0, step: 1, defaultValue: 100, unit: "m²/vozidlo" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 4.2, unit: "m³/m²" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.6 (provozní stavby; proxy)", kind: "number", min: 0, step: 1, defaultValue: 12255, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "HZS_COVERED_POPULATION", label: "Obyvatelé spádu pokrytí stávající jednotkou", kind: "number", min: 0, step: 1, unit: "obyvatel" },
    { code: "HZS_RESPONSE_MINUTES", label: "Ověřená doba příjezdu do řešeného území", kind: "number", min: 0, step: 0.1, unit: "min" },
    { code: "HZS_AVAILABLE_VEHICLES", label: "Ověřený počet dostupných vozidel / družstev", kind: "number", min: 0, step: 1, unit: "vozidel" },
    { code: "HZS_COVERAGE_SOURCE", label: "Zdroj / datum ověření požárního poplachového plánu", kind: "text", unit: "" }
  ],
  VYB_17A: [
    { code: "MODEL_POPULATION_PER_OFFICER", label: "Obyvatel na 1 policistu (scénář)", kind: "number", min: 1, step: 10, defaultValue: 300, unit: "obyvatel/policistu" },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Modelový maximální počet policistů na služebnu", kind: "number", min: 1, step: 1, defaultValue: 70, unit: "policistů/služebnu" },
    { code: "MODEL_BASE_AREA", label: "Minimální užitná plocha služebny", kind: "number", min: 0, step: 1, defaultValue: 250, unit: "m²/služebnu" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na jednoho policistu", kind: "number", min: 0, step: 0.1, defaultValue: 18, unit: "m²/policistu" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jednu služebnu", kind: "number", min: 0, step: 1, defaultValue: 2500, unit: "m²/služebnu" },
    { code: "MODEL_PAVED_PER_BASE", label: "Zpevněná plocha služebny", kind: "number", min: 0, step: 1, defaultValue: 500, unit: "m²/služebnu" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 4, unit: "m³/m²" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.6 (správa a administrativa; proxy)", kind: "number", min: 0, step: 1, defaultValue: 12255, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.23, unit: "%" },
    { code: "POLICE_RESPONSE_MINUTES", label: "Ověřená dojezdová doba hlídky do území", kind: "number", min: 0, step: 0.1, unit: "min" },
    { code: "POLICE_COVERAGE_SOURCE", label: "Zdroj / datum ověření pokrytí policie", kind: "text", unit: "" }
  ],
  VYB_17B: [
    { code: "MODEL_POPULATION_SHARE", label: "Podíl populace určený k ukrytí", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "Výchozí horní scénář počítá se všemi obyvateli ve spádu; místní plán může podíl upravit." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jednoho úkrytového modulu", kind: "number", min: 1, step: 1, defaultValue: 500, unit: "osob" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na osobu", kind: "number", min: 0, step: 0.1, defaultValue: 1.5, unit: "m²/osobu" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_CONSTRUCTION_HEIGHT", label: "Konstrukční výška podzemního úkrytu", kind: "number", min: 0.1, step: 0.1, defaultValue: 2.8, unit: "m" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na osobu", kind: "number", min: 0, step: 0.1, defaultValue: 1.5, unit: "m²/osobu" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.2, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "Sazba podzemní stavby (scénářový odhad, ne RTS)", kind: "number", min: 0, step: 100, defaultValue: 30000, unit: "Kč/m³", note: "Přiložené RTS 2026 neobsahuje samostatnou sazbu úkrytu. JKSO 814.2 oceňuje pozemní nádrže a 815.9 cenu za metr délky; nelze je vydávat za sazbu úkrytu v Kč/m³." },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.3, unit: "%" },
    { code: "SHELTER_PLANNED_DEMAND", label: "Počet osob určených k ukrytí podle ověřeného místního plánu (volitelné)", kind: "number", min: 0, step: 1, unit: "osob" },
    { code: "SHELTER_EXISTING_CAPACITY", label: "Ověřená kapacita stávajících úkrytů v dosahu 500 m", kind: "number", min: 0, step: 1, unit: "osob" },
    { code: "SHELTER_ACCESS_DISTANCE", label: "Skutečná pěší vzdálenost k započtenému úkrytu", kind: "number", min: 0, step: 1, unit: "m" },
    { code: "SHELTER_VERIFICATION_SOURCE", label: "Zdroj / datum ověření plánu, kapacity a trasy", kind: "text", unit: "" }
  ]
};
const TARGETED_PLAN_INPUT_CODES = new Set(
  Object.values(TARGETED_PLAN_INPUTS).flat().map((field) => field.code)
);
const TARGETED_FACILITY_ASSUMPTIONS = {
  VYB_01_J: { unitLabel: "míst", rateSource: "RTS 2026, JKSO 801.3 – budovy pro výuku a výchovu" },
  VYB_02C_1: { unitLabel: "žákovských míst", rateSource: "RTS 2026, JKSO 801.3 – budovy pro výuku a výchovu" },
  VYB_02C_2: { unitLabel: "žákovských míst", rateSource: "RTS 2026, JKSO 801.4 – budovy pro vědu, kulturu a osvětu (cenový proxy)" },
  VYB_09A_4: { unitLabel: "souč. zátěže lékařské ordinace", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_09D: { unitLabel: "výjezdových vozů", rateSource: "RTS 2026, JKSO 801.6 – správní/provozní budovy (cenový proxy)" },
  VYB_16A: { unitLabel: "požárních vozidel", rateSource: "RTS 2026, JKSO 801.6 – správní/provozní budovy (cenový proxy)" },
  VYB_17A: { unitLabel: "policejních míst", rateSource: "RTS 2026, JKSO 801.6 – budovy pro řízení, správu a administrativu" },
  VYB_17B: { unitLabel: "osob", rateSource: "Scénářový odhad mimo RTS; samostatná sazba úkrytu v RTS 2026 není uvedena" }
};
const TARGETED_RESULT_METRIC_LABELS = {
  VYB_01_J: ["Poptávka nepokrytá ověřenou kapacitou", "Navrhovaný počet dětských skupin", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_02C_1: ["Odhadovaná poptávka po místech", "Poptávka po odečtu ověřené volné kapacity", "Počet modelových zařízení", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_02C_2: ["Odhadovaná poptávka po místech", "Poptávka po odečtu ověřené volné kapacity", "Počet modelových zařízení", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_09A_4: ["Poptávka po gynekologii k ověření / po odečtu kapacity", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_09D: ["Modelová poptávka po výjezdových vozech", "Modelový počet výjezdových základen", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_16A: ["Doplňkový počet vozidel k prověření", "Modelový počet hasičských stanic", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_17A: ["Scénářový počet policistů k zajištění", "Modelová potřeba nových služeben po kontrole dojezdu", "Modelový počet policejních služeben", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_17B: ["Poptávka nepokrytá ověřenou kapacitou úkrytů", "Modelový počet úkrytových modulů", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"]
};
const isTargetedModelInput = (field) => field.code.startsWith("MODEL_")
  || field.code === "SHELTER_PLANNED_DEMAND";
const TARGETED_FIELD_COLORS = {
  MODEL_DEMOGRAPHIC_SHARE: "area",
  MODEL_ATTENDANCE_SHARE: "area",
  MODEL_POPULATION_PER_FTE: "area",
  MODEL_CAPACITY_PER_FACILITY: "area",
  MODEL_POPULATION_PER_VEHICLE: "area",
  MODEL_POPULATION_PER_OFFICER: "area",
  MODEL_POPULATION_SHARE: "area",
  MODEL_GROUP_CAPACITY: "area",
  MODEL_MIN_GROUP_CAPACITY: "area",
  MODEL_AREA_PER_UNIT: "area",
  MODEL_BASE_AREA: "area",
  MODEL_SITE_PER_UNIT: "site",
  MODEL_SITE_AREA_PER_UNIT: "site",
  MODEL_PAVED_SHARE: "paved",
  MODEL_PAVED_PER_BASE: "paved",
  MODEL_PAVED_PER_UNIT: "paved",
  MODEL_FLOOR_FACTOR: "floor",
  MODEL_VOLUME_PER_FLOOR_M2: "volume",
  MODEL_CONSTRUCTION_HEIGHT: "volume",
  MODEL_COST_PER_M3: "rts",
  MODEL_COST_SURCHARGE: "surcharge",
  ITEM_CONSTRUCTION_SYSTEM: "rts"
};
const HEALTH_SPECIALTIES = {
  VYB_16: [{ profession: "Lékárna / výdej hrazených léčiv", group: "Ověřit podle služby", limitMinutes: null }],
  VYB_17: [{ profession: "Praktický lékař pro dospělé", group: "Základní ambulantní péče", limitMinutes: 35 }],
  VYB_18: [{ profession: "Praktický lékař pro děti a dorost", group: "Základní ambulantní péče", limitMinutes: 35 }],
  VYB_19: [{ profession: "Zubní lékař", group: "Ověřit podle služby", limitMinutes: null }],
  VYB_20: [{ profession: "Gynekologie / sdružené ambulantní pracoviště", group: "Ověřit podle odbornosti", limitMinutes: null }],
  VYB_09A_4: [{ profession: "Gynekologie a porodnictví", group: "Ověřit podle druhu péče", limitMinutes: null }],
  VYB_09B_1: [
    { profession: "Chirurgie", group: "2", limitMinutes: 45 },
    { profession: "Interní lékařství", group: "2", limitMinutes: 45 },
    { profession: "Otorinolaryngologie (ORL)", group: "2", limitMinutes: 45 },
    { profession: "Oftalmologie", group: "2", limitMinutes: 45 }
  ],
  VYB_09B_2: [
    { profession: "Kardiologie", group: "3", limitMinutes: 60 },
    { profession: "Dermatovenerologie", group: "3", limitMinutes: 60 },
    { profession: "Neurologie dospělých", group: "Ověřit konkrétní službu", limitMinutes: null }
  ]
};
const HEALTHCARE_POPULATION_PER_FTE = {
  VYB_17: { value: 1800, label: "VPL" },
  VYB_18: { value: 4500, label: "PLDD" },
  VYB_19: { value: 2100, label: "stomatologie" },
  VYB_09A_4: { value: 14000, label: "gynekologie" }
};
const PARKING_SHARE_ASSUMPTIONS = {
  VYB_15A_3: { defaultPercent: 2.5, basis: "pracovní scénář sešitu; normativní zdroj pro 2,5 % nebyl doložen" },
  VYB_15B_1: { defaultPercent: 2, basis: "pracovní scénář sešitu; normativní zdroj pro 2 % nebyl doložen" },
  VYB_15B_2: { defaultPercent: 2.4, basis: "pracovní scénář sešitu; podíl 2,4 % obyvatel je orientační, nikoli zákonný požadavek" }
};
const CATEGORY_COLORS = ["sage", "blue", "amber", "rose", "violet", "teal", "slate"];
const ESTIMATE_FIELD_COLORS = {
  ITEM_CONSTRUCTION_SYSTEM: "rts",
  ITEM_USABLE_AREA_PER_UNIT: "area",
  ITEM_TOTAL_FLOOR_FACTOR: "floor",
  ITEM_VOLUME_PER_AREA: "volume",
  ITEM_SITE_AREA_PER_UNIT: "site",
  ITEM_PAVED_AREA_PER_UNIT: "paved",
  ITEM_RTS_RATE: "rts",
  ITEM_RTS_RATE_BASIS: "rts",
  ITEM_RTS_FIXED_COST: "fixed",
  ITEM_COST_SURCHARGE: "surcharge",
  K_CELKOVA_PLOCHA: "floor",
  K_OBESTAVENY: "volume",
  K_PLOCHA_AREALU: "site",
  K_ZASTAVENI: "site",
  K_ZPEVNENI: "paved",
  K_CENA_SKOLY: "rts",
  POLY_OBYV_NA_CENTRUM: "area",
  POLY_UZITNA_PLOCHA: "area",
  POLY_CENA_M3: "rts",
  POLY_PRIRAZKA: "surcharge",
  BYTY_PLOCHA_NA_BYT: "area",
  BYTY_OBESTAVENI_NA_M2: "volume",
  BYTY_CENA_M3: "rts",
  O_celk: "area",
  O_sidlo: "area",
  O_spad: "area",
  BJ_POCET_MANUAL: "area",
  PARKING_TOTAL_HPP: "floor",
  S_uzemi: "site",
  S_zastav: "site",
  k_odkup_obec: "area",
  m_marze_soc: "surcharge",
  PARKING_APARTMENT_SHARE_A: "area",
  PARKING_APARTMENT_SHARE_B: "area",
  PARKING_APARTMENT_SHARE_C: "area",
  PARKING_APARTMENT_SHARE_D: "area",
  PARKING_LARGE_APARTMENT_SHARE: "area",
  PARKING_SMALL_APARTMENT_AREA_M2: "area",
  PARKING_LARGE_APARTMENT_AREA_M2: "area",
  PARKING_FAMILY_HOME_AREA_M2: "area",
  PARKING_APARTMENT_RATE: "area",
  PARKING_FAMILY_HOME_RATE: "area",
  "HPP na byt": "floor",
  "Průměrná podlažnost": "floor",
  "Koeficient zastavění": "site",
  MAX_REZERVA: "area",
  PREKRYTI: "area"
};

function getEstimateColorForInput(field) {
  const label = String(field.label ?? "").toLocaleLowerCase("cs");
  const unit = String(field.unit ?? "").toLocaleLowerCase("cs");
  if (label.includes("zpevně")) return "paved";
  if (label.includes("pozem") || label.includes("areál") || label.includes("vzdálenost")
    || label.includes("trasa") || unit === "m") return "site";
  if (unit.includes("kč")) return "rts";
  if (unit.includes("m³")) return "volume";
  if (label.includes("podlaž") || label.includes("hpp") || unit.includes("podlaží")) return "floor";
  if (field.kind === "percent" && /náklad|rezerv|vrn/.test(label)) return "surcharge";
  if (unit.includes("m²") || unit.includes("obyvatel") || unit.includes("osob")) return "area";
  return ["number", "percent", "select"].includes(field.kind) ? "area" : null;
}

const state = {
  workbook: null,
  fileName: "",
  fileKey: "",
  items: [],
  rules: [],
  unlinkedRules: { legal: [], recommended: [], review: [] },
  inputs: [],
  categories: [],
  collapsedCategories: new Set(),
  collapsedLists: {},
  projectInputGroupsOpen: {},
  derivedDetailsOpen: false,
  selectedCategory: "",
  selectedItemKey: "",
  hideUnrealizedItems: false,
  search: "",
  primaryValues: {},
  secondaryValues: {},
  defaults: { primary: {}, secondary: {} },
  primaryLabels: {},
  healthcarePopulationPerFte: {},
  shapeConstants: {},
  peoplePerHome: 2.6,
  pedestrianDetourFactor: 1.3,
  capacityEdits: {},
  overrides: {},
  reviewChecks: {},
  standardComparisons: {},
  healthChecks: {},
  parkingShares: {},
  realizationSelections: {},
  itemEstimateValues: {},
  scenarioChanged: false,
  inputsChanged: false,
  persistenceError: false
};
const numberFormat = new Intl.NumberFormat("cs-CZ", { maximumFractionDigits: 2 });
const currencyFormat = new Intl.NumberFormat("cs-CZ", {
  style: "currency",
  currency: "CZK",
  maximumFractionDigits: 0
});

const elements = {
  workbookInput: document.querySelector("#workbook-input"),
  dropzone: document.querySelector("#dropzone"),
  excelExport: document.querySelector("#excel-export"),
  workbookName: document.querySelector("#workbook-name"),
  notice: document.querySelector("#notice"),
  itemTotal: document.querySelector("#item-total"),
  workbookStatus: document.querySelector("#workbook-status"),
  themeToggle: document.querySelector("#theme-toggle"),
  scenarioStatus: document.querySelector("#scenario-status"),
  resetInputs: document.querySelector("#reset-inputs"),
  projectInputGrid: document.querySelector("#project-input-grid"),
  derivedGrid: document.querySelector("#derived-grid"),
  secondaryInputGrid: document.querySelector("#secondary-input-grid"),
  search: document.querySelector("#search-input"),
  categoryList: document.querySelector("#category-list"),
  selectedCategoryTitle: document.querySelector("#selected-category-title"),
  visibleCount: document.querySelector("#visible-count"),
  hideUnrealizedItems: document.querySelector("#hide-unrealized-items"),
  legalCount: document.querySelector("#legal-count"),
  recommendedCount: document.querySelector("#recommended-count"),
  legalList: document.querySelector("#legal-list"),
  recommendedList: document.querySelector("#recommended-list"),
  infrastructureCount: document.querySelector("#infrastructure-count"),
  infrastructureList: document.querySelector("#infrastructure-list"),
  reviewCount: document.querySelector("#review-count"),
  reviewList: document.querySelector("#review-list"),
  requirementColumns: document.querySelector(".requirement-columns"),
  integrityWarning: document.querySelector("#integrity-warning")
};

function cellValue(row, index) {
  return index < 0 ? null : row[index] ?? null;
}

function renderIntegrityWarning() {
  const unlinkedCount = Object.values(state.unlinkedRules).reduce((total, rules) => total + rules.length, 0);
  const ambiguousIds = state.ambiguousModelIds || [];
  const messages = [];
  if (ambiguousIds.length) {
    const unresolvedIds = ambiguousIds.filter((id) => id !== "VYB_19");
    if (unresolvedIds.length) messages.push(`Další duplicitní kódy vybavenosti: ${unresolvedIds.join(", ")}.`);
  }
  if (unlinkedCount) {
    messages.push(`${unlinkedCount} pravidel nemá v sešitu přesné jednoznačné ID položky modelu. „Nepřiřazené ID“ neznamená, že potřeba s položkou věcně nesouvisí; dokud se nepotvrdí mapování, pravidlo se nepoužije k automatickému ocenění.`);
  }
  elements.integrityWarning.textContent = messages.join(" ");
  elements.integrityWarning.hidden = messages.length === 0;
}

function recordsFromSheet(sheet, headerKey, includeSourceRow = false) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  const headerRowIndex = rows.findIndex((row) => row.includes(headerKey));
  if (headerRowIndex < 0) throw new Error(`V listu „${sheet.name || "?"}“ chybí sloupec „${headerKey}“.`);
  const headers = rows[headerRowIndex].map((value) => String(value ?? "").trim());
  return rows.slice(headerRowIndex + 1)
    .map((row, index) => ({
      ...Object.fromEntries(headers.map((header, column) => [header, cellValue(row, column)])),
      ...(includeSourceRow ? { __sourceRow: headerRowIndex + index + 1 } : {})
    }))
    .filter((record) => Object.values(record).some((value) => value !== null && value !== ""));
}

function worksheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
}

function resolveRuleItems(equipmentIds, ruleId, items) {
  const resolved = [];
  const unresolvedIds = [];
  for (const id of equipmentIds) {
    const candidates = items.filter((item) => item.id === id);
    if (candidates.length === 1) {
      resolved.push(candidates[0]);
    } else if (candidates.length > 1 && ruleId.startsWith("ZDR-")) {
      const healthItems = candidates.filter((item) => /zdravot|health/i.test(item.category));
      if (healthItems.length === 1) resolved.push(healthItems[0]);
      else unresolvedIds.push(id);
    } else if (candidates.length > 1 && ruleId.startsWith("TECH-")) {
      const infrastructureItems = candidates.filter((item) => /technick|infrastruktur/i.test(item.category));
      if (infrastructureItems.length === 1) resolved.push(infrastructureItems[0]);
      else unresolvedIds.push(id);
    } else {
      unresolvedIds.push(id);
    }
  }
  return { items: resolved, unresolvedIds };
}

function readClusterMetric(modelSheet, sourceRow, modelColumn, metricColumn, workbook) {
  const formula = modelSheet[XLSX.utils.encode_cell({ r: sourceRow, c: modelColumn })]?.f;
  const match = formula?.match(/(?:'([^']+)'|(Cluster_[A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)/);
  if (!match) return null;
  const sheetName = match[1] || match[2];
  const clusterSheet = workbook.Sheets[sheetName];
  if (!clusterSheet) return null;
  const cell = clusterSheet[`${metricColumn}${match[4]}`];
  return numericOrNull(cell?.v);
}

function readCellFormula(sheet, sourceRow, column, sheetName, workbook = null) {
  const address = XLSX.utils.encode_cell({ r: sourceRow, c: column });
  const cell = sheet[address];
  const directReference = cell?.f?.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]{1,3})\$?(\d+)$/i);
  return {
    address: `${sheetName}!${address}`,
    sourceAddress: directReference
      ? `${directReference[1] || directReference[2]}!${directReference[3].toUpperCase()}${directReference[4]}`
      : null,
    formula: cell?.f ? `=${cell.f}` : null,
    value: cell?.v ?? null,
    calculation: cell?.f && workbook
      ? buildSimpleFormulaCalculation(`=${cell.f}`, sheetName, workbook) : null
  };
}

function readRtsPricing(modelCost, volumeSource, workbook) {
  const sourceAddress = modelCost?.sourceAddress || volumeSource?.address;
  const match = sourceAddress?.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!\$?[A-Z]{1,3}\$?(\d+)$/);
  if (!match) return null;
  const sheetName = match[1] || match[2];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;
  const row = match[3];
  const rateHeader = String(sheet.L8?.v ?? "");
  const rateBasis = /Kč\s*\/\s*MJ/i.test(rateHeader) ? "unit"
    : /Kč\s*\/\s*m\s*[³3]/i.test(rateHeader) ? "volume" : null;
  const rate = numericOrNull(sheet[`L${row}`]?.v);
  if (!rateBasis || rate === null || rate <= 0) return null;
  const modelZrn = numericOrNull(sheet[`P${row}`]?.v);
  const modelSurcharge = numericOrNull(sheet[`Q${row}`]?.v);
  const modelTotal = numericOrNull(sheet[`R${row}`]?.v);
  const modelVolume = numericOrNull(sheet[`O${row}`]?.v);
  const fixedCost = modelZrn !== null && modelVolume !== null
    ? Math.max(0, Math.round(modelZrn - modelVolume * rate)) : 0;
  const surcharge = modelZrn > 0 && modelTotal !== null
    ? Math.max(0, Math.round((modelTotal / modelZrn - 1) * 10000) / 10000)
    : modelSurcharge !== null && modelSurcharge >= 0
      ? modelSurcharge > 2 && modelZrn > 0
        ? Math.round((modelSurcharge / modelZrn) * 10000) / 10000 : modelSurcharge
      : null;
  return {
    rate,
    rateBasis,
    modelBasisQuantity: modelVolume,
    fixedCost,
    surcharge,
    sourceAddress: `${sheetName}!L${row}`
  };
}

function workbookCellDescriptor(workbook, sheetName, address) {
  const sheet = workbook.Sheets[sheetName];
  const match = address.match(/^([A-Z]+)(\d+)$/i);
  if (!sheet || !match) return null;
  const [, column, rowText] = match;
  const row = Number(rowText);
  const value = numericOrNull(sheet[address]?.v);
  if (value === null) return null;

  if (sheetName === "Vstupy_a_KPI") {
    const unit = String(sheet[`E${row}`]?.v || "").trim();
    return {
      value,
      label: String(sheet[`C${row}`]?.v || sheet[`B${row}`]?.v || "vstup").trim(),
      displayValue: unit.includes("%") ? value * 100 : value,
      unit: unit.includes("%") ? "%" : unit
    };
  }

  const columnIndex = XLSX.utils.decode_col(column);
  let label = "";
  for (let headerRow = row - 1; headerRow > 0 && headerRow >= row - 15; headerRow -= 1) {
    const candidate = sheet[`${column}${headerRow}`]?.v;
    if (typeof candidate === "string" && candidate.trim() && !candidate.startsWith("=")) {
      label = candidate.trim();
      break;
    }
  }
  const adjacentUnit = XLSX.utils.encode_cell({ r: row - 1, c: columnIndex + 1 });
  const unitValue = sheet[adjacentUnit]?.v;
  return {
    value,
    label: label || "hodnota modelu",
    displayValue: value,
    unit: typeof unitValue === "string" && !unitValue.startsWith("=") ? unitValue.trim() : ""
  };
}

function buildSimpleFormulaCalculation(formula, sheetName, workbook, depth = 0) {
  if (!formula) return null;
  const reference = String.raw`(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?\d+`;
  const operand = String.raw`(?:${reference}|\d+(?:\.\d+)?)`;
  const expression = formula.replace(/^=/, "").trim();
  const directReference = expression.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]{1,3})\$?(\d+)$/i);
  if (directReference && depth < 4) {
    const referencedSheet = directReference[1] || directReference[2];
    const address = `${directReference[3].toUpperCase()}${directReference[4]}`;
    const referencedCell = workbook.Sheets[referencedSheet]?.[address];
    if (referencedCell?.f) {
      return buildSimpleFormulaCalculation(`=${referencedCell.f}`, referencedSheet, workbook, depth + 1);
    }
  }
  const roundFunction = expression.match(/^(ROUNDUP|ROUND)\s*\(/i)?.[1]?.toUpperCase() || null;
  const match = expression.match(
    new RegExp(String.raw`^(?:(?:ROUNDUP|ROUND)\(\s*)?(${operand})\s*([*+/-])\s*(${operand})(?:\s*,\s*0\s*\))?$`, "i")
  );
  if (!match) return null;
  const resolve = (expression) => {
    const refMatch = expression.match(/^(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]{1,3})\$?(\d+)$/i);
    if (!refMatch) {
      const value = numericOrNull(expression);
      return value === null ? null : { value, displayValue: value, label: "koeficient", unit: "" };
    }
    const referencedSheet = refMatch[1] || refMatch[2] || sheetName;
    const address = `${refMatch[3].toUpperCase()}${refMatch[4]}`;
    return workbookCellDescriptor(workbook, referencedSheet, address);
  };
  const left = resolve(match[1]);
  const right = resolve(match[3]);
  if (!left || !right) return null;
  const operations = {
    "*": (a, b) => a * b,
    "/": (a, b) => b === 0 ? null : a / b,
    "+": (a, b) => a + b,
    "-": (a, b) => a - b
  };
  const exactValue = operations[match[2]](left.value, right.value);
  if (exactValue === null || !Number.isFinite(exactValue)) return null;
  const value = roundFunction === "ROUNDUP" ? Math.ceil(exactValue)
    : roundFunction === "ROUND" ? Math.round(exactValue) : exactValue;
  return {
    left,
    operator: match[2],
    right,
    value,
    rounded: Boolean(roundFunction),
    roundingMode: roundFunction === "ROUNDUP" ? "up" : roundFunction === "ROUND" ? "nearest" : null
  };
}

function formatWorkbookCalculation(source, label, result, unit) {
  const calculation = source?.calculation;
  if (!calculation) return null;
  const formatOperand = (operand) => {
    const unitText = operand.unit ? ` ${operand.unit}` : ` (${operand.label})`;
    return `${formatNumber(operand.displayValue ?? operand.value, 4)}${unitText}`;
  };
  const operator = calculation.operator === "*" ? "×"
    : calculation.operator === "/" ? "÷"
      : calculation.operator === "+" ? "+" : "−";
  const rounding = calculation.roundingMode === "up" ? " (zaokrouhleno nahoru na celé jednotky)"
    : calculation.rounded ? " (zaokrouhleno na celé jednotky)" : "";
  return `${label} = ${formatOperand(calculation.left)} ${operator} ${formatOperand(calculation.right)}${rounding} = ${formatValue(result, unit)}.`;
}

function readClusterMetricDetails(modelSheet, sourceRow, modelColumn, metricColumn, workbook) {
  const modelCell = readCellFormula(modelSheet, sourceRow, modelColumn, "Model_Vybavenosti", workbook);
  const match = modelCell.formula?.match(/(?:'([^']+)'|(Cluster_[A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)/);
  if (!match) return { value: null, modelCell, sourceCell: null };
  const sheetName = match[1] || match[2];
  const clusterSheet = workbook.Sheets[sheetName];
  if (!clusterSheet) return { value: null, modelCell, sourceCell: null };
  const sourceAddress = `${metricColumn}${match[4]}`;
  const sourceCell = {
    address: `${sheetName}!${sourceAddress}`,
    formula: clusterSheet[sourceAddress]?.f ? `=${clusterSheet[sourceAddress].f}` : null,
    value: clusterSheet[sourceAddress]?.v ?? null,
    calculation: buildSimpleFormulaCalculation(
      clusterSheet[sourceAddress]?.f ? `=${clusterSheet[sourceAddress].f}` : null,
      sheetName,
      workbook
    )
  };
  return {
    value: numericOrNull(sourceCell.value),
    modelCell,
    sourceCell
  };
}

function readProjectEstimateDefault(field, workbook) {
  const cellValueAt = (sheetName, address) => {
    const value = workbook.Sheets[sheetName]?.[address]?.v;
    return numericOrNull(value);
  };
  if (field.code === "K_OBESTAVENY") {
    return cellValueAt("Cluster_Zdravotnictvi", "D93") ?? field.defaultValue;
  }
  if (field.code === "K_PLOCHA_AREALU") {
    const site = cellValueAt("Cluster_Zdravotnictvi", "L93");
    const usable = cellValueAt("Cluster_Zdravotnictvi", "L92");
    return site !== null && usable > 0 ? Math.round((site / usable) * 10000) / 10000 : field.defaultValue;
  }
  if (field.code === "K_ZASTAVENI") {
    const floors = cellValueAt("Model_Vybavenosti", "R34");
    return floors > 0 ? 1 / floors : field.defaultValue;
  }
  if (field.code === "K_ZPEVNENI") {
    return cellValueAt("Model_Vybavenosti", "AD34") ?? field.defaultValue;
  }
  if (field.code === "POLY_OBYV_NA_CENTRUM") {
    const formula = workbook.Sheets.Cluster_Zdravotnictvi?.E13?.f || "";
    const threshold = formula.match(/D19\s*>=\s*(\d+(?:\.\d+)?)/i)?.[1];
    return numericOrNull(threshold) ?? field.defaultValue;
  }
  if (field.code === "POLY_UZITNA_PLOCHA") {
    return cellValueAt("Cluster_Zdravotnictvi", "L92") ?? field.defaultValue;
  }
  if (field.code === "POLY_CENA_M3") {
    return cellValueAt("Cluster_Zdravotnictvi", "L83") ?? field.defaultValue;
  }
  if (field.code === "POLY_PRIRAZKA") {
    const vrn = cellValueAt("Konstanty_a_Ciselniky", "C16");
    const reserve = cellValueAt("Konstanty_a_Ciselniky", "C17");
    return vrn !== null && reserve !== null ? vrn + reserve : field.defaultValue;
  }
  if (field.code === "BYTY_PLOCHA_NA_BYT") {
    const formula = workbook.Sheets.Cluster_Socialni_a_Bydleni?.H13?.f || "";
    const coefficient = formula.match(/E13\s*\*\s*(\d+(?:\.\d+)?)/i)?.[1];
    return numericOrNull(coefficient) ?? field.defaultValue;
  }
  if (field.code === "BYTY_OBESTAVENI_NA_M2") {
    const formula = workbook.Sheets.Cluster_Socialni_a_Bydleni?.J13?.f || "";
    const coefficient = formula.match(/H13\s*\*\s*(\d+(?:\.\d+)?)/i)?.[1];
    return numericOrNull(coefficient) ?? field.defaultValue;
  }
  if (field.code === "BYTY_CENA_M3") {
    return cellValueAt("Cluster_Socialni_a_Bydleni", "L13") ?? field.defaultValue;
  }
  return field.defaultValue;
}

function readEmbeddedProjectScenario(workbook) {
  const projectSheet = workbook.SheetNames
    .map((name) => workbook.Sheets[name])
    .find((sheet) => sheet?.A2?.v === "Formát" && sheet?.B2?.v === PROJECT_SHEET_MARKER);
  if (!projectSheet) return null;
  const rows = XLSX.utils.sheet_to_json(projectSheet, { header: 1, defval: null, raw: true });
  const chunks = rows.filter((row) => row[0] === "JSON_PART")
    .sort((left, right) => Number(left[1]) - Number(right[1]));
  if (!chunks.length || chunks.some((row, index) => Number(row[1]) !== index + 1 || typeof row[2] !== "string")) {
    throw new Error("List projektu je poškozený nebo neobsahuje všechna data scénáře.");
  }
  try {
    const data = JSON.parse(chunks.map((row) => row[2]).join(""));
    if (!data || typeof data !== "object" || !data.primaryValues || !data.realizationSelections) {
      throw new Error("Export neobsahuje platný scénář projektu.");
    }
    return data;
  } catch (error) {
    throw new Error(`Scénář projektu v sešitu se nepodařilo načíst: ${error.message}`);
  }
}

function readWorkbook(file) {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: "array", cellFormula: true, cellStyles: true });
    const missingSheets = REQUIRED_SHEETS.filter((name) => !workbook.Sheets[name]);
    if (missingSheets.length) throw new Error(`Sešit neobsahuje povinné listy: ${missingSheets.join(", ")}.`);
    const projectScenario = readEmbeddedProjectScenario(workbook);

    const modelSheet = workbook.Sheets.Model_Vybavenosti;
    const model = recordsFromSheet(modelSheet, MODEL_HEADERS.id, true);
    const modelRows = worksheetRows(modelSheet);
    const modelHeaders = modelRows.find((row) => row.includes(MODEL_HEADERS.id)) || [];
    const modelColumns = Object.fromEntries(modelHeaders.map((header, index) => [header, index]));
    const energyHeaders = modelRows.find((row) => row.includes(MODEL_HEADERS.annualTotalEnergy)) || [];
    const energyColumns = Object.fromEntries(energyHeaders.map((header, index) => [header, index]));
    const rules = recordsFromSheet(workbook.Sheets.Pravidla_a_vzorce, RULE_HEADERS.id);
    const inputs = recordsFromSheet(workbook.Sheets.Slovnik_vstupu, INPUT_HEADERS.id);
    let category = "Ostatní";
    const items = [];
    for (const record of model) {
      const id = String(record[MODEL_HEADERS.id] ?? "").trim();
      if (/^CLUSTER\b/i.test(id)) {
        category = id.replace(/^CLUSTER\s*\d*\s*:\s*/i, "").trim() || id;
        continue;
      }
      if (!/^VYB_/i.test(id)) continue;
      const name = String(record[MODEL_HEADERS.name] ?? "").trim();
      if (!name) continue;
      const sourceRow = record.__sourceRow;
      const formulaFor = (header) => {
        const column = modelColumns[header] ?? energyColumns[header];
        return column === undefined
          ? null : readCellFormula(modelSheet, sourceRow, column, "Model_Vybavenosti", workbook);
      };
      const energyValueFor = (header) => {
        const column = energyColumns[header];
        if (column === undefined) return null;
        const address = XLSX.utils.encode_cell({ r: sourceRow, c: column });
        return modelSheet[address]?.v ?? null;
      };
      const usableAreaDetails = readClusterMetricDetails(modelSheet, sourceRow, 15, "H", workbook);
      const parcelAreaDetails = readClusterMetricDetails(modelSheet, sourceRow, 16, "I", workbook);
      const volumeDetails = readClusterMetricDetails(modelSheet, sourceRow, 15, "J", workbook);
      const pavedAreaDetails = readClusterMetricDetails(
        modelSheet, sourceRow, modelColumns[MODEL_HEADERS.pavedArea], "AD", workbook
      );
      const costFormula = formulaFor(MODEL_HEADERS.cost);
      const rtsPricing = readRtsPricing(costFormula, volumeDetails.sourceCell, workbook);
      items.push({
        key: `${id}#${items.length + 1}`,
        id,
        name,
        category: INFRASTRUCTURE_ITEM_IDS.has(id) ? TECHNICAL_INFRASTRUCTURE_CATEGORY : category,
        unit: record[MODEL_HEADERS.unit],
        demand: numericOrNull(record[MODEL_HEADERS.demand]),
        capacityCount: numericOrNull(record[MODEL_HEADERS.capacityCount]),
        spatialCount: numericOrNull(record[MODEL_HEADERS.spatialCount]),
        proposed: numericOrNull(record[MODEL_HEADERS.proposed]),
        modelProposed: numericOrNull(record[MODEL_HEADERS.proposed]),
        areaPerPoint: numericOrNull(record[MODEL_HEADERS.areaPerPoint]),
        unitPrice: numericOrNull(record[MODEL_HEADERS.unitPrice]),
        cost: numericOrNull(record[MODEL_HEADERS.cost]),
        usableArea: usableAreaDetails.value,
        energyReferenceArea: numericOrNull(energyValueFor(MODEL_HEADERS.energyReferenceArea)),
        annualTotalEnergy: numericOrNull(energyValueFor(MODEL_HEADERS.annualTotalEnergy)),
        annualElectricity: numericOrNull(energyValueFor(MODEL_HEADERS.annualElectricity)),
        peakElectricalLoad: numericOrNull(energyValueFor(MODEL_HEADERS.peakElectricalLoad)),
        parcelArea: parcelAreaDetails.value,
        volume: volumeDetails.value,
        rtsPricing,
        sourceRow,
        sourceFormulas: {
          demand: formulaFor(MODEL_HEADERS.demand),
          capacityCount: formulaFor(MODEL_HEADERS.capacityCount),
          spatialCount: formulaFor(MODEL_HEADERS.spatialCount),
          proposed: formulaFor(MODEL_HEADERS.proposed),
          areaPerPoint: formulaFor(MODEL_HEADERS.areaPerPoint),
          unitPrice: formulaFor(MODEL_HEADERS.unitPrice),
          cost: costFormula,
          pavedArea: formulaFor(MODEL_HEADERS.pavedArea),
          annualTotalEnergy: formulaFor(MODEL_HEADERS.annualTotalEnergy),
          annualElectricity: formulaFor(MODEL_HEADERS.annualElectricity),
          peakElectricalLoad: formulaFor(MODEL_HEADERS.peakElectricalLoad)
        },
        metricSources: {
          usableArea: usableAreaDetails.sourceCell,
          parcelArea: parcelAreaDetails.sourceCell,
          volume: volumeDetails.sourceCell,
          pavedArea: pavedAreaDetails.sourceCell
        },
        coverage: record[MODEL_HEADERS.coverage],
        legalRules: [],
        recommendedRules: [],
        otherRules: [],
        manualCoverageCheck: REVIEW_ONLY_ITEM_IDS.has(id)
      });
      const item = items.at(-1);
      item.pavedArea = pavedAreaDetails.value ?? numericOrNull(record[MODEL_HEADERS.pavedArea]);
      if (CAPACITY_QUANTITY_IDS.has(id) && item.demand !== null) {
        if (item.modelProposed !== item.demand) item.sourceModelQuantity = item.modelProposed;
        item.proposed = item.demand;
        if (item.modelProposed !== item.demand) item.quantityBasis = "kapacitní poptávka";
      }
      if (id === "VYB_04A" && item.demand !== null) {
        item.capacityDemand = item.demand;
      }
    }
    items.push({
      key: `${HEATING_BOILER_ID}#1`,
      id: HEATING_BOILER_ID,
      name: "Kotelna pro ústřední vytápění",
      category: TECHNICAL_INFRASTRUCTURE_CATEGORY,
      unit: "kW tepelného výkonu",
      demand: null,
      capacityCount: null,
      spatialCount: null,
      proposed: 0,
      modelProposed: 0,
      areaPerPoint: null,
      unitPrice: null,
      cost: null,
      usableArea: null,
      energyReferenceArea: null,
      annualTotalEnergy: null,
      annualElectricity: null,
      peakElectricalLoad: null,
      parcelArea: null,
      volume: null,
      rtsPricing: null,
      sourceRow: null,
      sourceFormulas: {},
      metricSources: {},
      coverage: "Vloženo uživatelem — bez zákonného podkladu",
      legalRules: [],
      recommendedRules: [],
      otherRules: [],
      manualCoverageCheck: false,
      userAdded: true
    });
    if (!items.length) throw new Error("V listu Model_Vybavenosti nebyly nalezeny položky VYB_.");

    const rulesByEquipment = new Map();
    const unlinkedRules = { legal: [], recommended: [], review: [] };
    const itemCounts = new Map();
    items.forEach((item) => itemCounts.set(item.id, (itemCounts.get(item.id) || 0) + 1));
    const ambiguousModelIds = [...itemCounts.entries()].filter(([, count]) => count > 1).map(([id]) => id);
    for (const record of rules) {
      if (!record[RULE_HEADERS.id]) continue;
      const ruleId = String(record[RULE_HEADERS.id]).trim();
      const mappedIds = RULE_MODEL_LINKS[ruleId];
      const equipmentIds = mappedIds || String(record[RULE_HEADERS.equipmentId] ?? "")
        .split(/[,;/|]+/).map((value) => value.trim()).filter(Boolean);
      const { items: matchingItems, unresolvedIds } = resolveRuleItems(equipmentIds, ruleId, items);
      if (!matchingItems.length) {
        const group = isLegalRule(record) ? "legal" : isRecommendedRule(record) ? "recommended" : "review";
        const ambiguousIds = unresolvedIds.filter((id) => itemCounts.get(id) > 1);
        const issue = !equipmentIds.length
          ? "Záznam nemá uvedené ID vybavenosti."
          : ambiguousIds.length
            ? `ID ${ambiguousIds.join(", ")} se v modelu opakuje; vazba by nebyla jednoznačná.`
            : `ID ${equipmentIds.join(", ")} nemá jednoznačnou shodu v modelu.`;
        unlinkedRules[group].push({ ...record, linkIssue: issue });
      } else if (unresolvedIds.length) {
        const group = isLegalRule(record) ? "legal" : isRecommendedRule(record) ? "recommended" : "review";
        const ambiguousIds = unresolvedIds.filter((id) => itemCounts.get(id) > 1);
        const issue = ambiguousIds.length
          ? `Část ID (${ambiguousIds.join(", ")}) se v modelu opakuje; ostatní vazby jsou zobrazeny samostatně.`
          : `Část ID (${unresolvedIds.join(", ")}) nemá jednoznačnou shodu; ostatní vazby jsou zobrazeny samostatně.`;
        unlinkedRules[group].push({ ...record, linkIssue: issue });
      }
      const matchingIds = [...new Set(matchingItems.map((item) => item.id))];
      const linkedRecord = mappedIds && matchingItems.length
        ? { ...record, appLinkedIds: matchingIds, appMapping: true }
        : record;
      for (const item of matchingItems) {
        const group = rulesByEquipment.get(item.key) || [];
        group.push(linkedRecord);
        rulesByEquipment.set(item.key, group);
      }
    }
    for (const item of items) {
      for (const rule of rulesByEquipment.get(item.key) || []) {
        if (isLegalRule(rule)) item.legalRules.push(rule);
        else if (isRecommendedRule(rule)) item.recommendedRules.push(rule);
        else item.otherRules.push(rule);
      }
    }

    const inputRows = worksheetRows(workbook.Sheets.Vstupy_a_KPI);
    const primaryDefaults = {};
    const primaryLabels = {};
    primaryDefaults[PROJECT_LOCATION_INPUT.code] = PROJECT_LOCATION_INPUT.defaultValue;
    primaryLabels[PROJECT_LOCATION_INPUT.code] = PROJECT_LOCATION_INPUT.label;
    for (const field of PRIMARY_INPUTS) {
      const row = field.row ? inputRows[field.row - 1] : null;
      primaryDefaults[field.code] = field.defaultValue ?? cellValue(row, 3);
      primaryLabels[field.code] = field.label || cellValue(row, 2);
    }
    primaryDefaults.O_sidlo = PRIMARY_INPUTS.find((field) => field.code === "O_sidlo").defaultValue;
    primaryLabels.O_celk = PRIMARY_INPUTS.find((field) => field.code === "O_celk").label;
    primaryLabels.O_sidlo = PRIMARY_INPUTS.find((field) => field.code === "O_sidlo").label;
    primaryLabels.O_spad = PRIMARY_INPUTS.find((field) => field.code === "O_spad").label;
    primaryLabels.KOD_TVAR = PRIMARY_INPUTS.find((field) => field.code === "KOD_TVAR").label;
    const secondaryDefaults = {};
    for (const field of SECONDARY_INPUTS) {
      secondaryDefaults[field.code] = field.row
        ? cellValue(inputRows[field.row - 1], 7)
        : field.defaultValue;
    }
    secondaryDefaults[HALL_COVERAGE_INPUT.code] = cellValue(inputRows[HALL_COVERAGE_INPUT.row - 1], 7);
    for (const field of [...PROJECT_ESTIMATE_INPUTS, ...HOUSING_CONTRIBUTION_INPUTS]) {
      secondaryDefaults[field.code] = readProjectEstimateDefault(field, workbook);
    }
    secondaryDefaults[COVERAGE_OVERLAP_INPUT.code] = COVERAGE_OVERLAP_INPUT.defaultValue;
    const healthcarePopulationPerFte = {
      VYB_17: numericOrNull(workbook.Sheets.Cluster_Zdravotnictvi?.D10?.v),
      VYB_18: numericOrNull(workbook.Sheets.Cluster_Zdravotnictvi?.D11?.v),
      VYB_19: numericOrNull(workbook.Sheets.Cluster_Zdravotnictvi?.D12?.v)
    };
    const constantsRows = worksheetRows(workbook.Sheets.Konstanty_a_Ciselniky);
    const peoplePerHome = numericOrNull(cellValue(constantsRows[3], 2)) || 2.6;
    const pedestrianDetourFactor = numericOrNull(cellValue(constantsRows[4], 2)) || 1.3;
    const shapeConstants = {};
    for (let row = 19; row <= 23; row += 1) {
      const record = constantsRows[row - 1];
      const code = String(cellValue(record, 6) ?? "");
      if (!code) continue;
      shapeConstants[code] = {
        form: cellValue(record, 7),
        air: numericOrNull(cellValue(record, 8)),
        walk: numericOrNull(cellValue(record, 9)),
        network: numericOrNull(cellValue(record, 10))
      };
    }

    return {
      workbook,
      projectScenario,
      fileName: file.name,
      fileKey: `${file.name}:${file.size}:${file.lastModified}`,
      items,
      rules: rules.filter((record) => record[RULE_HEADERS.id]),
      unlinkedRules,
      ambiguousModelIds,
      inputs: inputs.filter((record) => record[INPUT_HEADERS.id]),
      categories: [...new Set(items.map((item) => item.category))],
      defaults: { primary: primaryDefaults, secondary: secondaryDefaults },
      primaryLabels,
      primaryValues: { ...primaryDefaults },
      secondaryValues: { ...secondaryDefaults },
      healthcarePopulationPerFte,
      peoplePerHome,
      pedestrianDetourFactor,
      shapeConstants
    };
  });
}

function numericOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function isLegalRule(rule) {
  const id = String(rule[RULE_HEADERS.id] ?? "").trim();
  if (id === "MOB-MET-002") return false;
  if (["CULT-LAW-001", "SOC-LEG-001"].includes(id)) return false;
  const area = `${rule[RULE_HEADERS.area] ?? ""} ${rule[RULE_HEADERS.type] ?? ""}`.toLocaleLowerCase("cs");
  return /zákonn|právn|legislativ|vyhlášk|zákonem|\bzávazn/.test(area);
}

function isRecommendedRule(rule) {
  if (String(rule[RULE_HEADERS.id] ?? "").trim() === "MOB-MET-002") return true;
  const area = `${rule[RULE_HEADERS.area] ?? ""} ${rule[RULE_HEADERS.type] ?? ""}`.toLocaleLowerCase("cs");
  return /metodik|doporuč|nezávazn|standard dostupnosti/.test(area);
}

function text(value, fallback = "Neuvedeno") {
  if (value === null || value === undefined || value === "") return fallback;
  return String(value);
}

function formatNumber(value, digits = 2) {
  if (typeof value !== "number" || !Number.isFinite(value)) return "—";
  return new Intl.NumberFormat("cs-CZ", { maximumFractionDigits: digits }).format(value);
}

function formatValue(value, unit = "") {
  if (value === null || value === undefined || value === "") return "—";
  return `${typeof value === "number" ? numberFormat.format(value) : value}${unit ? ` ${unit}` : ""}`;
}

function czechCountUnit(value, singular, few, plural) {
  if (value === 1) return singular;
  if (value >= 2 && value <= 4) return few;
  return plural;
}

function formatCost(value) {
  return typeof value === "number" && Number.isFinite(value) ? currencyFormat.format(value) : "Není stanoveno";
}

function getProjectViewState() {
  const collapsedLists = {};
  for (const [name, list] of [
    ["legal", elements.legalList],
    ["infrastructure", elements.infrastructureList],
    ["recommended", elements.recommendedList],
    ["review", elements.reviewList]
  ]) {
    const button = elements.requirementColumns.querySelector(`[data-collapse-list="${name}"]`);
    collapsedLists[name] = button
      ? button.getAttribute("aria-expanded") !== "true"
      : Boolean(list.hidden);
  }
  return {
    hideUnrealizedItems: elements.hideUnrealizedItems.checked,
    collapsedCategories: [...state.collapsedCategories],
    collapsedLists,
    projectInputGroupsOpen: { ...state.projectInputGroupsOpen },
    derivedDetailsOpen: document.querySelector(".derived-details")?.open ?? false,
    theme: document.documentElement.dataset.theme === "night" ? "night" : "day",
    selectedCategory: state.selectedCategory,
    selectedItemKey: state.selectedItemKey,
    search: elements.search.value
  };
}

function getProjectScenarioData() {
  return {
    schemaVersion: 7,
    primaryValues: state.primaryValues,
    secondaryValues: state.secondaryValues,
    capacityEdits: state.capacityEdits,
    overrides: state.overrides,
    reviewChecks: state.reviewChecks,
    standardComparisons: state.standardComparisons,
    healthChecks: state.healthChecks,
    parkingShares: state.parkingShares,
    realizationSelections: state.realizationSelections,
    itemEstimateValues: state.itemEstimateValues,
    scenarioChanged: state.scenarioChanged,
    inputsChanged: state.inputsChanged,
    ...getProjectViewState()
  };
}

function applyProjectScenarioData(data) {
  const savedPrimaryValues = { ...data.primaryValues };
  if ((data.schemaVersion || 1) < 2) delete savedPrimaryValues.O_sidlo;
  if ((data.schemaVersion || 1) < 4 && savedPrimaryValues.O_sidlo == null) {
    savedPrimaryValues.O_sidlo = 1500;
  }
  state.primaryValues = { ...state.primaryValues, ...savedPrimaryValues };
  state.secondaryValues = { ...state.secondaryValues, ...data.secondaryValues };
  state.capacityEdits = data.capacityEdits || {};
  state.overrides = data.overrides || {};
  state.reviewChecks = data.reviewChecks || {};
  state.standardComparisons = data.standardComparisons || {};
  state.healthChecks = data.healthChecks || {};
  state.parkingShares = data.parkingShares || {};
  state.realizationSelections = data.realizationSelections || {};
  state.itemEstimateValues = data.itemEstimateValues || {};
  state.scenarioChanged = Boolean(data.scenarioChanged);
  state.inputsChanged = Boolean(data.inputsChanged);
  state.hideUnrealizedItems = Boolean(data.hideUnrealizedItems);
  state.collapsedCategories = new Set(Array.isArray(data.collapsedCategories) ? data.collapsedCategories : []);
  state.collapsedLists = data.collapsedLists && typeof data.collapsedLists === "object"
    ? data.collapsedLists : {};
  state.projectInputGroupsOpen = data.projectInputGroupsOpen && typeof data.projectInputGroupsOpen === "object"
    ? data.projectInputGroupsOpen : {};
  state.derivedDetailsOpen = Boolean(data.derivedDetailsOpen);
  if (data.theme === "day" || data.theme === "night") applyTheme(data.theme);
  state.selectedCategory = state.categories.includes(data.selectedCategory) ? data.selectedCategory : "";
  state.selectedItemKey = state.items.some((item) => item.key === data.selectedItemKey)
    ? data.selectedItemKey : "";
  if (state.selectedItemKey) {
    state.selectedCategory = state.items.find((item) => item.key === state.selectedItemKey)?.category || "";
  }
  state.search = String(data.search || "");
}

function createElement(tag, className = "", content) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (content !== undefined && content !== null) element.textContent = String(content);
  return element;
}

function inputDescriptor(code) {
  return state.inputs.find((input) => input[INPUT_HEADERS.code] === code)
    || state.inputs.find((input) => String(input[INPUT_HEADERS.id]) === code)
    || { [INPUT_HEADERS.name]: code, [INPUT_HEADERS.unit]: "", [INPUT_HEADERS.verification]: "" };
}

function readScenario(embeddedScenario = null) {
  if (embeddedScenario) {
    applyProjectScenarioData(embeddedScenario);
    return null;
  }
  if (!state.fileKey) return null;
  try {
    const saved = localStorage.getItem(`territory-scenario:${state.fileKey}`);
    if (!saved) return null;
    const data = JSON.parse(saved);
    if (data && typeof data === "object") {
      applyProjectScenarioData(data);
    }
    return null;
  } catch (error) {
    state.persistenceError = true;
    return `Uložený scénář se nepodařilo načíst: ${error.message}`;
  }
}

function saveScenario(markScenarioChanged = true) {
  if (!state.fileKey) return;
  try {
    const data = getProjectScenarioData();
    data.scenarioChanged = markScenarioChanged || state.scenarioChanged;
    localStorage.setItem(`territory-scenario:${state.fileKey}`, JSON.stringify(data));
    if (markScenarioChanged) state.scenarioChanged = true;
    state.persistenceError = false;
  } catch (error) {
    state.scenarioChanged = false;
    state.persistenceError = true;
    showNotice(`Změnu se nepodařilo lokálně uložit: ${error.message}`, true);
  }
}

function getPrimary(code) {
  return state.primaryValues[code];
}

function getSecondary(code) {
  return state.secondaryValues[code];
}

function getResidentialMixAssumptions() {
  const territoryType = getResidentParkingTerritoryType();
  const apartmentShareCode = `PARKING_APARTMENT_SHARE_${territoryType}`;
  const apartmentShare = numericOrNull(getPrimary(apartmentShareCode)) ?? 0;
  const smallApartmentArea = numericOrNull(getPrimary("PARKING_SMALL_APARTMENT_AREA_M2")) ?? 85;
  const largeShare = numericOrNull(getPrimary("PARKING_LARGE_APARTMENT_SHARE")) ?? 0.1;
  const largeApartmentArea = numericOrNull(getPrimary("PARKING_LARGE_APARTMENT_AREA_M2")) ?? 150;
  const familyHomeArea = numericOrNull(getPrimary("PARKING_FAMILY_HOME_AREA_M2")) ?? 140;
  const averageApartmentArea = smallApartmentArea * (1 - largeShare) + largeApartmentArea * largeShare;
  const averageDwellingArea = apartmentShare * averageApartmentArea + (1 - apartmentShare) * familyHomeArea;
  return {
    territoryType,
    apartmentShare: Math.min(1, Math.max(0, apartmentShare)),
    largeShare: Math.min(1, Math.max(0, largeShare)),
    smallApartmentArea,
    largeApartmentArea,
    averageApartmentArea,
    familyHomeArea,
    averageDwellingArea
  };
}

function derivedValues() {
  const population = numericOrNull(getPrimary("O_celk"));
  const currentAreaPopulation = numericOrNull(getPrimary("O_sidlo"));
  const catchment = numericOrNull(getPrimary("O_spad"));
  const land = numericOrNull(getPrimary("S_uzemi"));
  const housingLand = numericOrNull(getPrimary("S_zastav"));
  const floorPerHome = numericOrNull(getSecondary("HPP na byt"));
  const floors = numericOrNull(getSecondary("Průměrná podlažnost"));
  const coverage = numericOrNull(getSecondary("Koeficient zastavění"));
  const shape = state.shapeConstants[getPrimary("KOD_TVAR")];
  const totalHousingHpp = numericOrNull(getPrimary("PARKING_TOTAL_HPP"));
  const residentialMix = getResidentialMixAssumptions();
  const enteredHouseholds = numericOrNull(getPrimary("BJ_POCET_MANUAL"));
  const households = enteredHouseholds !== null ? enteredHouseholds
    : totalHousingHpp > 0 && residentialMix.averageDwellingArea > 0
      ? Math.floor(totalHousingHpp / residentialMix.averageDwellingArea + 0.5)
    : population !== null ? Math.trunc(population / state.peoplePerHome) : null;
  const projectedAreaPopulation = currentAreaPopulation !== null && population !== null
    ? currentAreaPopulation + population : null;
  const result = {
    BJ_pocet: households,
    HPP_byty: totalHousingHpp > 0 ? totalHousingHpp
      : households !== null && floorPerHome !== null
        ? households * floorPerHome : null,
    O_uzemi_cil: projectedAreaPopulation,
    O_sidlo: projectedAreaPopulation,
    O_prah: projectedAreaPopulation !== null && catchment !== null ? projectedAreaPopulation + catchment : null,
    rho_hruba: population !== null && land > 0 ? population / land : null,
    rho_cista: population !== null && housingLand > 0 ? population / housingLand : null,
    k_tvar: shape?.air ?? null,
    k_pesi: shape?.walk ?? null,
    c_tvar: shape?.network ?? null
  };
  const rootArea = land !== null && land >= 0 ? Math.sqrt(land * 10000) : null;
  result.R_max = rootArea !== null && result.k_tvar !== null ? result.k_tvar * rootArea : null;
  result.L_max = rootArea !== null && result.k_pesi !== null ? result.k_pesi * rootArea : null;
  result.areaWhole = result.HPP_byty !== null && floors > 0 ? Math.round(result.HPP_byty / floors) : null;
  result.parcelWhole = result.HPP_byty !== null && floors > 0 && coverage > 0
    ? Math.round(result.HPP_byty / floors / coverage) : null;
  return result;
}

function formatPercent(value) {
  const number = numericOrNull(value);
  return number === null ? "" : String(number * 100);
}

function createInputControl(field, value, secondary = false, idSuffix = "", inputScope = "", itemKey = "") {
  const wrapper = createElement("div", "project-field");
  const color = field.color || ESTIMATE_FIELD_COLORS[field.code]
    || TARGETED_FIELD_COLORS[field.code] || getEstimateColorForInput(field);
  if (color) wrapper.dataset.estimateColor = color;
  const controlId = `${secondary ? "secondary" : "primary"}-${field.code}${idSuffix ? `-${idSuffix}` : ""}`;
  const title = field.label || (secondary ? field.code
    : state.primaryLabels[field.code] || inputDescriptor(field.code)[INPUT_HEADERS.name]);
  const titleLabel = createElement("label", "project-field-label", text(title, field.code));
  titleLabel.htmlFor = controlId;
  if (!secondary && field.code === "TYP_UZEMI") {
    const labelRow = createElement("div", "project-field-label-row");
    labelRow.append(titleLabel);
    const help = createElement("a", "input-help-link", "?");
    help.href = TERRITORY_TYPE_HELP_URL;
    help.target = "_blank";
    help.rel = "noopener noreferrer";
    help.title = "Jak vybrat kategorii: metodika TAČR je v přiloženém PDF. Klasifikaci nelze určit jen podle počtu obyvatel.";
    help.setAttribute("aria-label", help.title);
    labelRow.append(help);
    wrapper.append(labelRow);
  } else {
    wrapper.append(titleLabel);
  }
  let control;
  if (field.kind === "select") {
    control = document.createElement("select");
    field.options.forEach((option, index) => {
      const optionElement = new Option(field.labels?.[index] || option, option);
      control.append(optionElement);
    });
    control.value = value === null || value === undefined ? field.options[0] : String(value);
  } else {
    control = document.createElement("input");
    control.type = field.kind === "text" ? "text" : "number";
    const scale = field.kind === "percent" ? 100 : 1;
    if (field.kind !== "text") {
      control.min = String((field.min ?? 0) * scale);
      if (field.max !== undefined) control.max = String(field.max * scale);
      control.step = String((field.step ?? "any") === "any" ? "any" : field.step * scale);
      control.inputMode = field.kind === "percent" ? "decimal" : "numeric";
    }
  }
  control.id = controlId;
  control.dataset.inputKind = field.kind;
  control.dataset.inputCode = field.code;
  control.dataset.inputScope = inputScope || (secondary ? "secondary" : "primary");
  if (itemKey) control.dataset.itemKey = itemKey;
  if (field.kind === "percent") control.value = formatPercent(value);
  else control.value = value === null || value === undefined ? "" : String(value);
  if (value !== null && value !== undefined && value !== "") {
    control.dataset.initialValue = String(control.value);
  }
  wrapper.append(control);
  if (field.note) {
    wrapper.append(createElement("small", "project-field-note", field.note));
  }
  if (field.kind === "derived" && field.formula) {
    const formula = field.code === "BJ_pocet"
      ? numericOrNull(getPrimary("BJ_POCET_MANUAL")) !== null
        ? "BJ_POCET_MANUAL"
        : `INT(O_celk / ${formatNumber(state.peoplePerHome)})`
      : field.formula;
    wrapper.append(createElement("code", "project-field-formula", formula));
  }
  const unit = field.unit !== undefined
    ? field.unit : secondary ? "" : inputDescriptor(field.code)[INPUT_HEADERS.unit];
  const displayedUnit = field.kind === "percent" ? field.unit || "%"
    : field.code === "KOMPAKT" || /1\s*=\s*ano\s*\/\s*0\s*=\s*ne/i.test(String(unit ?? ""))
      || isDimensionlessUnit(unit) ? "" : unit;
  if (displayedUnit) {
    const valueRow = createElement("div", "project-field-value");
    valueRow.append(control, createElement("span", "project-field-unit", displayedUnit));
    wrapper.append(valueRow);
  } else {
    wrapper.append(control);
  }
  if (field.kind === "derived") control.readOnly = true;
  return wrapper;
}

function renderProjectInputs() {
  for (const group of elements.projectInputGrid.querySelectorAll("details[data-input-group]")) {
    state.projectInputGroupsOpen[group.dataset.inputGroup] = group.open;
  }
  elements.projectInputGrid.replaceChildren();
  const allFields = [...PRIMARY_INPUTS, ...SECONDARY_INPUTS];
  for (const groupDefinition of PROJECT_INPUT_GROUPS) {
    const isAlwaysOpen = groupDefinition.label === "Parametry projektu";
    const group = isAlwaysOpen
      ? createElement("section", "project-input-group project-input-group-pinned")
      : createElement("details", "project-input-group");
    group.dataset.inputGroup = groupDefinition.label;
    if (!isAlwaysOpen) group.open = Boolean(state.projectInputGroupsOpen[groupDefinition.label]);
    if (!isAlwaysOpen) {
      group.addEventListener("toggle", () => {
        state.projectInputGroupsOpen[groupDefinition.label] = group.open;
        saveScenario(false);
      });
    }
    const heading = document.createElement(isAlwaysOpen ? "h3" : "summary");
    heading.className = "project-input-group-heading";
    heading.append(createElement("span", "", groupDefinition.label));
    group.append(heading);
    const fields = createElement("div", "project-input-group-fields");
    for (const definition of groupDefinition.fields) {
      const field = definition.field || allFields.find((candidate) => candidate.code === definition.code);
      if (!field) continue;
      const value = definition.secondary ? getSecondary(field.code) : getPrimary(field.code);
      fields.append(createInputControl(field, value, Boolean(definition.secondary)));
    }
    group.append(fields);
    elements.projectInputGrid.append(group);
  }
  renderDerivedInputs();
  renderSecondaryInputs();
  const derivedDetails = document.querySelector(".derived-details");
  if (derivedDetails) derivedDetails.open = state.derivedDetailsOpen;
  elements.resetInputs.disabled = false;
  updateScenarioStatus();
}

function renderDerivedInputs() {
  const values = derivedValues();
  elements.derivedGrid.replaceChildren();
  for (const field of DERIVED_INPUTS) {
    const descriptor = inputDescriptor(field.code);
    const card = createElement("div", "derived-card");
    card.append(createElement("span", "derived-name", text(field.label || descriptor[INPUT_HEADERS.name], field.code)));
    const unit = field.unit ?? "";
    const value = values[field.code];
    const displayedUnit = isDimensionlessUnit(unit) ? "" : unit;
    card.append(createElement("strong", "derived-value", formatValue(value, displayedUnit)));
    card.append(createElement("code", "derived-formula", field.formula));
    elements.derivedGrid.append(card);
  }
}

function isDimensionlessUnit(unit) {
  return String(unit ?? "").toLocaleLowerCase("cs").includes("bezrozměrný");
}

function renderSecondaryInputs() {
  elements.secondaryInputGrid.replaceChildren();
  const heading = createElement("h3", "secondary-heading", "Další výpočtové předpoklady");
  elements.secondaryInputGrid.append(heading);
  const grid = createElement("div", "secondary-fields");
  const groupedCodes = new Set(PROJECT_INPUT_GROUPS.flatMap((group) =>
    group.fields.map((definition) => definition.code || definition.field?.code)));
  for (const field of SECONDARY_INPUTS.filter((candidate) => !groupedCodes.has(candidate.code))) {
    grid.append(createInputControl(field, getSecondary(field.code), true));
  }
  elements.secondaryInputGrid.append(grid);
}

function updateScenarioStatus() {
  const changed = state.scenarioChanged;
  elements.scenarioStatus.textContent = state.persistenceError
    ? "Změny nejsou uloženy"
    : changed
    ? "Scénář uložen v tomto prohlížeči"
    : "Výchozí hodnoty pracovního sešitu";
  elements.workbookStatus.textContent = changed ? "Scénář upraven" : "Sešit načten";
}

function handleInputChange(event) {
  const control = event.target.closest("[data-input-code]");
  if (!control) return;
  const disclosureState = captureRequirementDisclosureState();
  const inputScope = control.dataset.inputScope;
  const secondary = inputScope === "secondary";
  const kind = control.dataset.inputKind;
  let value = control.value;
  if (kind === "number" || kind === "percent") {
    if (value === "") {
      if (kind === "number" && Number(control.min) > 0) {
        control.setCustomValidity("Zadejte hodnotu v povoleném rozsahu.");
        control.reportValidity();
        return;
      }
      value = null;
    } else {
      value = Number(value);
      if (!Number.isFinite(value) || value < Number(control.min)
        || (control.dataset.inputCode === "BJ_POCET_MANUAL" && !Number.isInteger(value))
        || (control.max !== "" && value > Number(control.max))) {
        control.setCustomValidity("Hodnota je mimo povolený rozsah.");
        control.reportValidity();
        return;
      }
      if (kind === "percent") value /= 100;
    }
  }
  control.setCustomValidity("");
  if (inputScope === "itemEstimate") {
    const itemKey = control.dataset.itemKey;
    if (!state.itemEstimateValues[itemKey]) state.itemEstimateValues[itemKey] = {};
    state.itemEstimateValues[itemKey][control.dataset.inputCode] = value;
  } else {
    const target = secondary ? state.secondaryValues : state.primaryValues;
    target[control.dataset.inputCode] = value;
    state.inputsChanged = true;
  }
  saveScenario();
  renderProjectInputs();
  renderCategories();
  renderRequirements();
  renderIntegrityWarning();
  restoreRequirementDisclosureState(disclosureState);
}

function updateTargetedPlanInput(event) {
  const control = event.target.closest('[data-input-scope="itemEstimate"][data-input-code]');
  if (!control || !TARGETED_PLAN_INPUT_CODES.has(control.dataset.inputCode)) return false;
  const item = state.items.find((candidate) => candidate.key === control.dataset.itemKey);
  if (!item) return true;
  const itemField = (TARGETED_PLAN_INPUTS[item.id] || [])
    .find((field) => field.code === control.dataset.inputCode);
  if (!itemField) return true;
  const isIntegerInput = itemField.kind === "number" && itemField.step === 1;
  let value = control.value;
  if (["number", "percent"].includes(itemField.kind)) {
    const scale = itemField.kind === "percent" ? 100 : 1;
    if (value === "") {
      if (Number(control.min) > 0) {
        control.setCustomValidity("Zadejte hodnotu v povoleném rozsahu.");
        if (event.type === "change") control.reportValidity();
        return true;
      }
      value = null;
    } else {
      const displayedValue = Number(value);
      if (!Number.isFinite(displayedValue) || displayedValue < Number(control.min)
        || (control.max !== "" && displayedValue > Number(control.max))
        || isIntegerInput && !Number.isInteger(displayedValue)) {
        control.setCustomValidity(isIntegerInput
          ? "Zadejte celé nezáporné číslo v povoleném rozsahu."
          : "Zadejte hodnotu v povoleném rozsahu.");
        if (event.type === "change") control.reportValidity();
        return true;
      }
      value = displayedValue / scale;
    }
  }
  control.setCustomValidity("");
  if (!state.itemEstimateValues[item.key]) state.itemEstimateValues[item.key] = {};
  state.itemEstimateValues[item.key][control.dataset.inputCode] = value;
  saveScenario();

  const card = control.closest(".requirement-card");
  const kind = card?.dataset.kind;
  if (!card || !kind) return true;
  card.querySelector(".item-results")?.replaceWith(renderItemMetrics(item, kind));
  const estimate = getTargetedPlanEstimate(item);
  const verificationMetrics = card.querySelector(".targeted-verification-metrics");
  if (verificationMetrics) {
    verificationMetrics.replaceWith(renderTargetedVerificationMetrics(estimate));
  }
  const verificationMethods = card.querySelector(".targeted-verification-method");
  if (verificationMethods) {
    const updatedMethods = renderTargetedVerificationMethods(estimate, item);
    if (updatedMethods) {
      updatedMethods.open = verificationMethods.open;
      verificationMethods.replaceWith(updatedMethods);
    } else {
      verificationMethods.remove();
    }
  }
  const calculation = card.querySelector(".project-estimate-calculation");
  if (calculation) {
    const updatedCalculation = renderProjectEstimateCalculation(item, kind);
    const updatedLines = [...updatedCalculation.children]
      .filter((element) => element.classList.contains("calculation-equation")
        || element.classList.contains("calculation-source-note"));
    calculation.querySelectorAll(".calculation-equation, .calculation-source-note")
      .forEach((element) => element.remove());
    calculation.append(...updatedLines);
  }
  const cost = estimate.metrics.find((metric) => metric.label.includes("investice"))?.value;
  const quantitySummary = card.querySelector(".requirement-quantity");
  if (quantitySummary) quantitySummary.textContent = getTargetedPlanSummaryText(item, estimate);
  const investmentSummary = card.querySelector(".requirement-investment");
  if (investmentSummary) investmentSummary.textContent = cost === null || cost === undefined
    ? "bez samostatného ocenění" : formatCost(cost);
  updateScenarioStatus();
  return true;
}

function captureRequirementDisclosureState() {
  return [...document.querySelectorAll(".requirement-card")].map((card) => ({
    kind: card.dataset.kind,
    itemKey: card.dataset.itemKey,
    open: card.open,
    sections: [...card.querySelectorAll("details")].map((details) => details.open)
  }));
}

function restoreRequirementDisclosureState(disclosureState) {
  for (const saved of disclosureState) {
    const card = [...document.querySelectorAll(".requirement-card")]
      .find((candidate) => candidate.dataset.kind === saved.kind
        && candidate.dataset.itemKey === saved.itemKey);
    if (!card) continue;
    card.open = saved.open;
    card.querySelectorAll("details").forEach((details, index) => {
      details.open = Boolean(saved.sections[index]);
    });
  }
}

function handleInlineProjectInputChange(event) {
  if (updateTargetedPlanInput(event)) return;
  handleInputChange(event);
}

function categoryIndex(category) {
  const index = state.categories.indexOf(category);
  return index < 0 ? 0 : index % CATEGORY_COLORS.length;
}

function categoryShape(category) {
  return ["circle", "square", "triangle", "diamond", "hexagon", "pentagon", "star"][
    categoryIndex(category) % 7
  ];
}

function matchingItems() {
  const query = state.search.toLocaleLowerCase("cs").trim();
  return state.items.filter((item) => {
    if (state.selectedCategory && item.category !== state.selectedCategory) return false;
    if (state.selectedItemKey && item.key !== state.selectedItemKey) return false;
    if (!query) return true;
    const ruleText = [...item.legalRules, ...item.recommendedRules].map((rule) =>
      [rule[RULE_HEADERS.id], rule[RULE_HEADERS.area], rule[RULE_HEADERS.source], rule[RULE_HEADERS.description]]
        .join(" ")).join(" ");
    return `${item.id} ${item.name} ${item.category} ${ruleText}`.toLocaleLowerCase("cs").includes(query);
  });
}

function renderCategories() {
  elements.categoryList.replaceChildren();
  const realizationPlan = getRealizationPlan();
  const realizationSummary = createElement("div", "realization-summary");
  realizationSummary.append(createElement("strong", "",
    `Přímo vybráno: ${realizationPlan.directKeys.size}`));
  realizationSummary.append(createElement("span", "",
    "Žádná položka ani infrastruktura se nezařazuje automaticky."));
  realizationSummary.append(createElement("span", "realization-legend",
    "Barevný tvar označuje skupinu; plné zvýraznění znamená výběr k realizaci."));
  elements.categoryList.append(realizationSummary);
  const allButton = createElement("button", `category-button${!state.selectedCategory && !state.selectedItemKey ? " is-active" : ""}`);
  allButton.type = "button";
  allButton.dataset.category = "";
  allButton.append(createElement("span", "category-swatch all-swatch"));
  allButton.append(createElement("span", "category-name", "Všechny skupiny"));
  allButton.append(createElement("span", "category-count", state.items.length));
  elements.categoryList.append(allButton);
  state.categories.forEach((category, index) => {
    const categoryItems = state.items.filter((item) => item.category === category);
    const group = createElement("section", `category-group category-group-${CATEGORY_COLORS[index % CATEGORY_COLORS.length]}`);
    const button = createElement("button", `category-button category-${index % CATEGORY_COLORS.length}${state.selectedCategory === category && !state.selectedItemKey ? " is-active" : ""}`);
    button.type = "button";
    button.dataset.category = category;
    button.append(createElement("span", "category-swatch"));
    button.append(createElement("span", "category-name", category));
    button.append(createElement("span", "category-count", categoryItems.length));
    const collapsed = state.collapsedCategories.has(category);
    const listId = `category-items-${index}`;
    const header = createElement("div", "category-group-header");
    header.append(button);
    const toggle = createElement("button", "collapse-button category-collapse-button");
    toggle.type = "button";
    toggle.dataset.collapseCategory = category;
    toggle.setAttribute("aria-expanded", String(!collapsed));
    toggle.setAttribute("aria-controls", listId);
    toggle.setAttribute("aria-label", `${collapsed ? "Rozbalit" : "Sbalit"} skupinu ${category}`);
    header.append(toggle);
    group.append(header);
    const items = createElement("div", "category-equipment-list");
    items.id = listId;
    items.hidden = collapsed;
    categoryItems.forEach((item) => {
      const selected = realizationPlan.selectedKeys.has(item.key);
      const shape = categoryShape(category);
      const selectionLabel = createElement("label",
        `realization-select-label category-${categoryIndex(category)} shape-${shape}`);
      const selection = document.createElement("input");
      selection.type = "checkbox";
      selection.checked = selected;
      selection.value = selected ? "1" : "0";
      selection.dataset.realizationItemKey = item.key;
      selection.setAttribute("aria-label", `${selected ? "Zrušit výběr" : "Vybrat"}: ${displayCatalogItemName(item)}`);
      selection.title = "Označit tuto položku (hodnota 1), nebo její výběr zrušit (hodnota 0).";
      selectionLabel.append(selection);
      const itemButton = createElement("button",
        `category-equipment-button${state.selectedItemKey === item.key ? " is-active" : ""}${selected ? " is-realization-selected" : ""}`,
        displayCatalogItemName(item));
      itemButton.type = "button";
      itemButton.dataset.itemKey = item.key;
      if (item.userAdded) {
        const userMark = createElement("span", "user-inserted-mark", "Uživatel");
        userMark.title = "Položka byla vložena uživatelem; nemá zákonný podklad.";
        itemButton.append(userMark);
      }
      const row = createElement("div", "category-equipment-row");
      row.append(selectionLabel, itemButton);
      items.append(row);
    });
    group.append(items);
    elements.categoryList.append(group);
  });
}

function getRealizationPlan() {
  const directItems = state.items.filter((item) =>
    Number(state.realizationSelections[item.key]) === 1);
  const directKeys = new Set(directItems.map((item) => item.key));
  return {
    directKeys,
    selectedKeys: new Set(directKeys)
  };
}

function updateRealizationSelection(event) {
  const input = event.target.closest("[data-realization-item-key]");
  if (!input) return;
  const item = state.items.find((candidate) => candidate.key === input.dataset.realizationItemKey);
  if (!item || input.disabled) return;
  setRealizationSelection(item, input.checked);
}

function setRealizationSelection(item, selected) {
  state.realizationSelections[item.key] = selected ? 1 : 0;
  for (const control of document.querySelectorAll(
    `[data-realization-item-key="${CSS.escape(item.key)}"]`
  )) {
    if (control instanceof HTMLInputElement) {
      control.checked = selected;
      control.value = selected ? "1" : "0";
      control.setAttribute("aria-label",
        `${selected ? "Zrušit výběr" : "Vybrat"}: ${displayCatalogItemName(item)}`);
      continue;
    }
    control.classList.toggle("is-selected", selected);
    control.setAttribute("aria-pressed", String(selected));
    control.setAttribute("aria-label",
      `${selected ? "Zrušit výběr" : "Vybrat"} k realizaci: ${displayCatalogItemName(item)}`);
    control.title = selected ? "Odebrat z realizace" : "Vybrat k realizaci";
  }
  document.querySelectorAll(".category-equipment-button[data-item-key]").forEach((button) => {
    if (button.dataset.itemKey === item.key) button.classList.toggle("is-realization-selected", selected);
  });
  document.querySelectorAll(".requirement-card[data-item-key]").forEach((card) => {
    if (card.dataset.itemKey === item.key) card.classList.toggle("is-realization-selected", selected);
  });
  const count = elements.categoryList.querySelector(".realization-summary strong");
  if (count) {
    count.textContent = `Přímo vybráno: ${getRealizationPlan().directKeys.size}`;
  }
  if (state.hideUnrealizedItems) updateRequirementVisibility();
  saveScenario();
  updateScenarioStatus();
}

function getPlannedQuantity(item, kind = "") {
  const override = state.overrides[item.key];
  if (override && Number.isFinite(override.quantity) && override.note?.trim()) return override.quantity;
  if (isResidentParkingItem(item)) {
    const legal = getResidentParkingLegalEstimate(item);
    const recommended = getResidentParkingRecommendedEstimate(item);
    if (item.id === RESIDENT_PARKING_ID) return kind === "legal" ? legal?.requiredStalls ?? null : null;
    if (recommended === null) return null;
    if (item.id === RESIDENT_PARKING_LAW_ID) {
      return kind === "recommended" ? recommended.additionalStalls : null;
    }
    if (item.id === RESIDENT_PARKING_RESERVE_ID) {
      return kind === "recommended" ? recommended.reserveStalls : null;
    }
  }
  const standardTarget = getAppliedStandardTarget(item);
  if (standardTarget !== null) return standardTarget;
  return item.proposed;
}

function getResidentParkingLegalEstimate(item) {
  if (!isResidentParkingItem(item)) return null;
  const saved = state.itemEstimateValues[item.parkingSourceKey || item.key] || {};
  const manualFloorArea = numericOrNull(saved.PARKING_LEGAL_AREA_M2);
  const minimumFactor = numericOrNull(saved.PARKING_LEGAL_FACTOR_MIN) ?? 1;
  const units = numericOrNull(derivedValues().BJ_pocet);
  const mix = getResidentialMixAssumptions();
  const apartmentUnits = units === null ? null : units * mix.apartmentShare;
  const familyHomeUnits = units === null ? null : units - apartmentUnits;
  const smallApartmentUnits = apartmentUnits === null ? null : apartmentUnits * (1 - mix.largeShare);
  const largeApartmentUnits = apartmentUnits === null ? null : apartmentUnits * mix.largeShare;
  const modeledApartmentArea = smallApartmentUnits === null || largeApartmentUnits === null ? null
    : smallApartmentUnits * mix.smallApartmentArea + largeApartmentUnits * mix.largeApartmentArea;
  const modeledFamilyHomeArea = familyHomeUnits === null ? null : familyHomeUnits * mix.familyHomeArea;
  const modeledArea = modeledApartmentArea === null || modeledFamilyHomeArea === null
    ? null : modeledApartmentArea + modeledFamilyHomeArea;
  const suppliedHpp = numericOrNull(getPrimary("PARKING_TOTAL_HPP"));
  const estimatedFloorArea = suppliedHpp > 0 ? suppliedHpp : modeledArea;
  const floorArea = manualFloorArea ?? estimatedFloorArea;
  const areaScale = floorArea !== null && modeledArea > 0 ? floorArea / modeledArea : 1;
  const apartmentArea = floorArea === null ? null
    : modeledApartmentArea !== null && modeledArea > 0
      ? modeledApartmentArea * areaScale : floorArea * mix.apartmentShare;
  const familyHomeArea = floorArea === null ? null
    : modeledFamilyHomeArea !== null && modeledArea > 0
      ? modeledFamilyHomeArea * areaScale : floorArea - apartmentArea;
  const buyoutShare = numericOrNull(getPrimary("k_odkup_obec")) ?? 0;
  const accessibleQualified = getPrimary("PARKING_1B_QUALIFIED") !== "NE";
  const accessibleShare = accessibleQualified ? Math.min(1, Math.max(0, buyoutShare)) : 0;
  const affordableRentalArea = apartmentArea === null ? null : apartmentArea * accessibleShare;
  const effectiveOrdinaryArea = floorArea === null || affordableRentalArea === null
    ? null : floorArea - affordableRentalArea;
  const ordinaryStalls = effectiveOrdinaryArea === null
    ? null : effectiveOrdinaryArea / 120 * minimumFactor;
  const affordableRentalStalls = affordableRentalArea === null
    ? null : affordableRentalArea / 240 * minimumFactor;
  const baseStalls = ordinaryStalls === null || affordableRentalStalls === null
    ? null : ordinaryStalls + affordableRentalStalls;
  const requiredStalls = baseStalls === null ? null : Math.floor(baseStalls + 0.5);
  return {
    floorArea,
    manualFloorArea,
    estimatedFloorArea,
    units,
    apartmentUnits,
    smallApartmentUnits,
    largeApartmentUnits,
    familyHomeUnits,
    apartmentArea,
    familyHomeArea,
    smallApartmentArea: mix.smallApartmentArea,
    largeApartmentArea: mix.largeApartmentArea,
    familyHomeAreaPerUnit: mix.familyHomeArea,
    largeApartmentShare: mix.largeShare,
    affordableRentalArea,
    accessibleShare,
    accessibleQualified,
    buyoutShare,
    ordinaryStalls,
    affordableRentalStalls,
    minimumFactor,
    ordinaryAreaPerStall: 120,
    affordableAreaPerStall: 240,
    baseStalls,
    requiredStalls
  };
}

function getResidentParkingReservePercent() {
  const reserveItem = state.items.find((item) => item.id === RESIDENT_PARKING_RESERVE_ID);
  const saved = reserveItem
    ? state.itemEstimateValues[reserveItem.key]?.PARKING_RESERVE_PERCENT
    : null;
  const value = numericOrNull(saved) ?? numericOrNull(getPrimary("PARKING_RESERVE_PERCENT"));
  return value === null ? 0.1 : Math.min(1, Math.max(0, value));
}

function getResidentParkingTerritoryType() {
  return String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase() ?? "A";
}

function getResidentParkingRecommendedEstimate(item) {
  if (!isResidentParkingItem(item)) return null;
  const legal = getResidentParkingLegalEstimate(item);
  const mix = getResidentialMixAssumptions();
  const apartmentUnits = legal?.apartmentUnits;
  const familyHomeUnits = legal?.familyHomeUnits;
  if (legal?.requiredStalls === null || apartmentUnits === null || familyHomeUnits === null) return null;
  const apartmentRate = numericOrNull(getPrimary("PARKING_APARTMENT_RATE")) ?? 1.1;
  const familyHomeRate = numericOrNull(getPrimary("PARKING_FAMILY_HOME_RATE")) ?? 2;
  const modeledApartmentStalls = Math.floor(apartmentUnits * apartmentRate + 0.5);
  const modeledFamilyHomeStalls = Math.floor(familyHomeUnits * familyHomeRate + 0.5);
  const modeledStalls = modeledApartmentStalls + modeledFamilyHomeStalls;
  const additionalStalls = Math.max(0, modeledStalls - legal.requiredStalls);
  const combinedBaseStalls = legal.requiredStalls + additionalStalls;
  const reservePercent = getResidentParkingReservePercent();
  const reserveStalls = Math.ceil(combinedBaseStalls * reservePercent);
  return {
    population: numericOrNull(getPrimary("O_celk")),
    units: legal.units,
    apartmentUnits,
    familyHomeUnits,
    territoryType: getResidentParkingTerritoryType(),
    apartmentShare: mix.apartmentShare,
    apartmentRate,
    familyHomeRate,
    modeledStalls,
    modeledApartmentStalls,
    modeledFamilyHomeStalls,
    baseStalls: combinedBaseStalls,
    additionalStalls,
    combinedBaseStalls,
    selectedBaseStalls: combinedBaseStalls,
    reservePercent,
    reserveStalls,
    totalWithReserve: combinedBaseStalls + reserveStalls
  };
}

function getResidentParkingRecommendedExtra(item) {
  const legal = getResidentParkingLegalEstimate(item);
  const recommended = getResidentParkingRecommendedEstimate(item);
  if (legal?.requiredStalls === null || recommended === null) return null;
  return recommended.additionalStalls;
}

function getSchoolRecommendationPercent(item) {
  const saved = numericOrNull(state.standardComparisons[item.key]?.recommended);
  if (saved !== null && saved >= 0 && saved <= 100) return saved;
  const settlementType = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  return ({ A: 10, B: 10, C: 15, D: 20 })[settlementType] ?? 10;
}

function getRecommendedIncrease(item) {
  if (isResidentParkingItem(item)) return null;
  const value = numericOrNull(state.standardComparisons[item.key]?.recommended);
  return value !== null && value >= 0 ? value : null;
}

function getStandardBaseQuantity(item) {
  return numericOrNull(usesDemandBasedCapacity(item) ? item.demand : item.proposed);
}

function isWholeNumberUnit(unit) {
  return /(dět|žák|lůž|bytů|stání|míst|zařízen|škola|osob|tříd|objekt|stanic)/i.test(String(unit ?? ""));
}

function getReserveLimit() {
  const value = numericOrNull(getSecondary("MAX_REZERVA"));
  return value === null ? 0.5 : Math.min(1, Math.max(0, value));
}

function getAllowedCapacityUse(item, available, demand) {
  const allowed = Math.min(available * getReserveLimit(), demand);
  const unit = String(getStandardComparisonUnit(item) ?? "").toLocaleLowerCase("cs");
  return /(dět|žák|lůž|bytů|stání|míst|zařízen|škola|osob)/.test(unit)
    ? Math.floor(allowed)
    : allowed;
}

function getVerifiedCapacity(item) {
  const capacity = getCapacityEdit(item);
  const quantity = numericOrNull(capacity.quantity);
  return quantity !== null && quantity >= 0 ? quantity : 0;
}

function getStandardComparisonUnit(item) {
  return usesDemandBasedCapacity(item) || isHealthcareStaffingItem(item)
    ? getCapacityUnit(item) : item.unit;
}

function getStandardComparisonResultUnit(item) {
  return usesDemandBasedCapacity(item) || isHealthcareStaffingItem(item)
    ? getCapacityResultUnit(item) : item.unit;
}

function formatStandardComparisonAmount(item, value) {
  let unit = getStandardComparisonResultUnit(item);
  if (Number.isInteger(value)) {
    if (item.id === "VYB_01") {
      unit = value === 1 ? "dítě" : value >= 2 && value <= 4 ? "děti" : "dětí";
    } else if (item.id === "VYB_02A" || item.id === "VYB_02B") {
      const grade = item.id === "VYB_02A" ? "1." : "2.";
      unit = value === 1 ? `žák ${grade} stupně`
        : value >= 2 && value <= 4 ? `žáci ${grade} stupně` : `žáků ${grade} stupně`;
    }
  }
  return `${formatNumber(value)} ${text(unit, "")}`;
}

function hasAppliedStandardTarget(item) {
  return getPrimary("REZIM_STD") === "STANDARD"
    && hasStandardComparison(item)
    && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
    && getRecommendedIncrease(item) !== null
    && getStandardBaseQuantity(item) !== null;
}

function getAppliedStandardTarget(item) {
  if (!hasAppliedStandardTarget(item)) return null;
  const comparison = getStandardComparison(item);
  return comparison.base + comparison.uplift;
}

function getRequirementUnit(item) {
  if (usesDemandBasedCapacity(item) || isHealthcareStaffingItem(item)) {
    return getCapacityResultUnit(item);
  }
  return hasAppliedStandardTarget(item) ? getStandardComparisonUnit(item) : item.unit;
}

function getSummaryRequirementQuantity(calc) {
  if (calc.usesDemandBasedCapacity && calc.capacityDemand !== null) {
    return calc.isCapacityVerified ? calc.remainingCapacity : calc.capacityDemand;
  }
  return calc.remaining;
}

function getCapacitySummaryText(item, calc) {
  return `poptávka ${formatStandardComparisonAmount(item, calc.capacityDemand)}; započteno ${formatStandardComparisonAmount(item, calc.verifiedCapacity)}`;
}

function getStandardMetricScale(item, kind, calc = adjustedResult(item)) {
  if (!hasStandardComparison(item) || SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) return 1;
  const comparison = getStandardComparison(item);
  if (comparison.uplift === null) return 1;
  if (comparison.base === null || comparison.base <= 0) return 1;
  if (calc.hasOverride) return 1;
  if (kind === "recommended" && getPrimary("REZIM_STD") !== "STANDARD") {
    return comparison.delta / comparison.base;
  }
  if (getPrimary("REZIM_STD") === "STANDARD" && hasAppliedStandardTarget(item)) {
    const remaining = calc.usesDemandBasedCapacity ? calc.remainingCapacity : calc.remaining;
    return remaining === null ? getAppliedStandardTarget(item) / comparison.base
      : remaining / comparison.base;
  }
  return 1;
}

function getStandardMetricValue(value, scale, unit = "") {
  if (value === null || value === undefined) return null;
  const scaled = value * scale;
  return isWholeNumberUnit(unit) ? Math.ceil(scaled) : scaled;
}

function sameMetricValue(left, right) {
  const leftNumber = numericOrNull(left);
  const rightNumber = numericOrNull(right);
  if (leftNumber === null || rightNumber === null) return leftNumber === rightNumber;
  return Math.abs(leftNumber - rightNumber) <= Math.max(1, Math.abs(leftNumber), Math.abs(rightNumber)) * 1e-9;
}

function getAggregateProjectSurfaceMetrics() {
  const realizationPlan = getRealizationPlan();
  let totalFloorArea = 0;
  let roofArea = 0;
  let pavedArea = 0;
  const residentialBase = derivedValues();
  totalFloorArea += numericOrNull(residentialBase.HPP_byty) ?? 0;
  roofArea += numericOrNull(residentialBase.areaWhole) ?? 0;
  for (const candidate of state.items) {
    if (!realizationPlan.selectedKeys.has(candidate.key)
      || TECHNICAL_ALTERNATIVE_IDS.has(candidate.id) || candidate.id === "VYB_07B") continue;
    const estimate = getProjectPlanningEstimate(candidate, "recommended");
    if (!estimate) continue;
    if (estimate.type === "school" || estimate.type === "polyclinic") {
      totalFloorArea += numericOrNull(estimate.totalFloorArea) ?? 0;
      roofArea += numericOrNull(estimate.footprintArea) ?? 0;
      pavedArea += numericOrNull(estimate.pavedArea) ?? 0;
    } else if (estimate.type === "targeted" && estimate.facilityPlan) {
      totalFloorArea += numericOrNull(estimate.facilityPlan.totalFloorArea) ?? 0;
      roofArea += numericOrNull(estimate.facilityPlan.footprintArea) ?? 0;
      pavedArea += numericOrNull(estimate.facilityPlan.pavedArea) ?? 0;
    } else if (estimate.type === "item") {
      const area = numericOrNull(estimate.totalFloorArea) ?? 0;
      const floorFactor = numericOrNull(estimate.totalFloorFactor) ?? 1;
      totalFloorArea += area;
      roofArea += floorFactor > 0 ? area / floorFactor : 0;
      pavedArea += numericOrNull(estimate.pavedArea) ?? 0;
    }
  }
  return { totalFloorArea, roofArea, pavedArea, runoffArea: roofArea + pavedArea };
}

function getAggregateProjectEnergyMetrics(item) {
  const realizationPlan = getRealizationPlan();
  let annualElectricity = 0;
  let peakElectricalLoad = 0;
  let modeledNonElectricEnergy = 0;
  let modeledEquipmentCount = 0;
  let fallbackFloorArea = 0;
  let fallbackHeatingArea = 0;
  let projectFloorArea = 0;
  const addEnergySource = (candidate, housingEstimate = null) => {
    const estimate = housingEstimate ? null : getProjectPlanningEstimate(candidate, "recommended");
    const usableArea = numericOrNull(housingEstimate?.portfolioArea)
      ?? numericOrNull(estimate?.usableArea)
      ?? numericOrNull(candidate.usableArea);
    const totalFloorArea = numericOrNull(housingEstimate?.portfolioArea)
      ?? numericOrNull(estimate?.totalFloorArea)
      ?? usableArea;
    if (totalFloorArea === null || totalFloorArea <= 0) return;
    projectFloorArea += totalFloorArea;
    const referenceArea = numericOrNull(candidate.energyReferenceArea);
    const annualTotal = numericOrNull(candidate.annualTotalEnergy);
    const itemElectricity = numericOrNull(candidate.annualElectricity);
    const itemPeak = numericOrNull(candidate.peakElectricalLoad);
    if (referenceArea > 0 && usableArea > 0 && annualTotal !== null
      && itemElectricity !== null && itemPeak !== null) {
      const scale = usableArea / referenceArea;
      annualElectricity += itemElectricity * scale;
      peakElectricalLoad += itemPeak * scale;
      if (!REALIZATION_PARKING_IDS.has(candidate.id)) {
        modeledNonElectricEnergy += Math.max(0, annualTotal - itemElectricity) * scale;
      }
      modeledEquipmentCount += 1;
    } else {
      fallbackFloorArea += totalFloorArea;
      if (!REALIZATION_PARKING_IDS.has(candidate.id)) fallbackHeatingArea += totalFloorArea;
    }
  };
  const housingItem = state.items.find((candidate) => candidate.id === "VYB_07B");
  const residentialFloorArea = numericOrNull(derivedValues().HPP_byty) ?? 0;
  if (housingItem && residentialFloorArea > 0) {
    addEnergySource(housingItem, { portfolioArea: residentialFloorArea });
  } else {
    projectFloorArea += residentialFloorArea;
    fallbackFloorArea += residentialFloorArea;
    fallbackHeatingArea += residentialFloorArea;
  }
  for (const candidate of state.items) {
    if (!realizationPlan.selectedKeys.has(candidate.key)
      || TECHNICAL_ALTERNATIVE_IDS.has(candidate.id)
      || REALIZATION_NON_BUILDING_IDS.has(candidate.id)) continue;
    addEnergySource(candidate);
  }
  const energyDensity = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_DENSITY_W_M2")) ?? 0;
  const peakToAverageRatio = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO")) ?? 4;
  const fallbackPeakElectricalLoad = fallbackFloorArea * energyDensity / 1000;
  const fallbackAnnualElectricity = peakToAverageRatio > 0
    ? fallbackPeakElectricalLoad * 8760 / peakToAverageRatio : 0;
  const fallbackHeatDensity = numericOrNull(getItemEstimateValue(item, "BOILER_HEAT_DENSITY_KWH_M2_Y")) ?? 0;
  return {
    annualElectricity: annualElectricity + fallbackAnnualElectricity,
    peakElectricalLoad: peakElectricalLoad + fallbackPeakElectricalLoad,
    modeledAnnualElectricity: annualElectricity,
    modeledPeakElectricalLoad: peakElectricalLoad,
    modeledNonElectricEnergy,
    modeledEquipmentCount,
    fallbackAnnualElectricity,
    fallbackPeakElectricalLoad,
    fallbackFloorArea,
    fallbackHeatDemand: fallbackHeatingArea * fallbackHeatDensity,
    fallbackHeatingArea,
    projectFloorArea
  };
}

function getItemEstimateDefaultQuantity(item) {
  if (item.id === "VYB_05B") {
    return getItemEstimateValue(item, HALL_COVERAGE_INPUT.code) === "Samostatná nová hala požadována"
      ? 1 : 0;
  }
  if (item.id === "VYB_10B_1") {
    const localPopulation = (numericOrNull(getPrimary("O_celk")) ?? 0)
      + (numericOrNull(getPrimary("O_sidlo")) ?? 0);
    return localPopulation > 1000 ? 1 : 0;
  }
  const parkingShare = PARKING_SHARE_ASSUMPTIONS[item.id];
  if (parkingShare) {
    const population = numericOrNull(getPrimary("O_celk"));
    return population === null ? 0 : Math.ceil(population * getParkingShare(item) / 100);
  }
  const planned = numericOrNull(getPlannedQuantity(item));
  if (planned !== null && planned >= 0) return planned;
  const quantity = numericOrNull(item.modelProposed);
  if (quantity !== null && quantity >= 0) return quantity;
  const demand = numericOrNull(item.demand);
  return demand !== null && demand >= 0 ? demand : 0;
}

function getItemEstimateQuantity(item) {
  const saved = numericOrNull(state.itemEstimateValues[item.key]?.ITEM_QUANTITY);
  return saved !== null && saved >= 0 ? saved : getItemEstimateDefaultQuantity(item);
}

function estimateComparableUnitRate(item) {
  const rates = state.items
    .filter((candidate) => candidate.key !== item.key
      && candidate.category === item.category
      && candidate.unit === item.unit
      && numericOrNull(candidate.unitPrice) > 0)
    .map((candidate) => numericOrNull(candidate.unitPrice))
    .sort((left, right) => left - right);
  if (!rates.length) return null;
  const middle = Math.floor(rates.length / 2);
  return rates.length % 2
    ? rates[middle]
    : (rates[middle - 1] + rates[middle]) / 2;
}

function getItemEstimateDefault(item, code) {
  const quantity = getItemEstimateDefaultQuantity(item);
  const scenario = ITEM_ESTIMATE_SCENARIOS[item.id];
  const perUnit = (value) => {
    const number = numericOrNull(value);
    return number !== null && number > 0 && quantity > 0 ? number / quantity : null;
  };
  switch (code) {
    case "ITEM_QUANTITY": return getItemEstimateDefaultQuantity(item);
    case "ITEM_USABLE_AREA_PER_UNIT":
      if (item.id === "VYB_15B_2") return 1.5;
      if (scenario) return scenario.usableArea;
      return perUnit(item.usableArea) ?? perUnit(item.areaPerPoint);
    case "ITEM_TOTAL_FLOOR_FACTOR":
      if (item.id === "VYB_15B_2") return 1;
      if (scenario) return scenario.floorFactor;
      return numericOrNull(item.volume) > 0 ? numericOrNull(getSecondary("K_CELKOVA_PLOCHA")) ?? 1 : null;
    case "ITEM_VOLUME_PER_AREA": {
      if (item.id === "VYB_15B_2") return 1;
      if (scenario) return scenario.volumeFactor;
      const sourceVolume = numericOrNull(item.volume);
      const sourceArea = numericOrNull(item.usableArea);
      const floorFactor = numericOrNull(getSecondary("K_CELKOVA_PLOCHA")) ?? 1;
      return sourceVolume !== null && sourceVolume > 0 && sourceArea > 0 && floorFactor > 0
        ? sourceVolume / (sourceArea * floorFactor)
        : null;
    }
    case "ITEM_SITE_AREA_PER_UNIT":
      if (item.id === "VYB_15B_2") return 1.2;
      if (scenario) return scenario.siteArea;
      return perUnit(item.parcelArea) ?? perUnit(item.areaPerPoint);
    case "ITEM_PAVED_AREA_PER_UNIT":
      if (item.id === "VYB_15B_2") return 1.5;
      if (scenario) return scenario.pavedArea;
      return perUnit(item.pavedArea);
    case "ITEM_RTS_RATE": {
      const constructionRate = getRtsConstructionRate(item);
      if (constructionRate !== undefined) return constructionRate;
      if (item.id === "VYB_15B_2") return 10100;
      if (item.id === "VYB_10B_1") return 13645;
      if (item.rtsPricing?.rate) return numericOrNull(item.rtsPricing.rate);
      if (numericOrNull(item.unitPrice) > 0) return numericOrNull(item.unitPrice);
      if (scenario) return 13645;
      const sourceQuantity = getItemEstimateDefaultQuantity(item);
      if (numericOrNull(item.cost) > 0 && sourceQuantity > 0) {
        return numericOrNull(item.cost) / sourceQuantity / 1.23;
      }
      return estimateComparableUnitRate(item);
    }
    case "ITEM_RTS_RATE_BASIS":
      if (getRtsBuildingGroup(item)) return "volume";
      if (item.id === "VYB_15B_2") return "unit";
      if (item.id === "VYB_10B_1") return "volume";
      if (scenario) return "volume";
      return item.rtsPricing?.rateBasis || "unit";
    case "ITEM_RTS_FIXED_COST": return numericOrNull(item.rtsPricing?.fixedCost) ?? 0;
    case "ITEM_COST_SURCHARGE": {
      if (item.id === "VYB_10B_1") return 0.23;
      const sourceSurcharge = numericOrNull(item.rtsPricing?.surcharge);
      if (sourceSurcharge !== null) return sourceSurcharge;
      const legacyFactor = numericOrNull(getSecondary("K_CENA_SKOLY"));
      return SCHOOL_CAPACITY_ITEM_IDS.has(item.id) && legacyFactor !== null
        ? Math.max(0, legacyFactor - 1) : 0.23;
    }
    default: return null;
  }
}

function getAlternativeEstimateDefault(code) {
  switch (code) {
    case "ALT_RAINFALL_DEPTH_MM": return 15;
    case "ALT_RUNOFF_COEFFICIENT": return 0.9;
    case "ALT_RAINWATER_RATE": return 13600;
    case "ALT_WATER_LITERS_PER_PERSON_DAY": return 87;
    case "ALT_PEAK_DAY_FACTOR": return 2;
    case "ALT_WASTEWATER_RETURN": return 0.9;
    case "ALT_NETWORK_LENGTH_M": return 1500;
    case "ALT_WATER_PIPE_RATE": return 5231.2;
    case "ALT_WATER_METER_CHAMBER_COUNT": return 1;
    case "ALT_WATER_METER_CHAMBER_RATE": return 150000;
    case "ALT_SEWER_PIPE_RATE": return 4814.53;
    case "ALT_SEWER_MANHOLE_SPACING_M": return 50;
    case "ALT_SEWER_MANHOLE_RATE": return 25000;
    case "ALT_ENERGY_DENSITY_W_M2": return 50;
    case "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO": return 4;
    case "BOILER_HEAT_DENSITY_KWH_M2_Y": return 60;
    case "BOILER_HEATING_SHARE": return 1;
    case "BOILER_FULL_LOAD_HOURS": return 1800;
    case "BOILER_DESIGN_RESERVE": return 1.2;
    case "BOILER_UNIT_CAPACITY_KW": return 500;
    case "BOILER_UNIT_COST": return null;
    case "ALT_POWER_CABLE_RATE": return 2569.6;
    case "ALT_TRANSFORMER_CAPACITY_KW": return 2500;
    case "ALT_TRANSFORMER_RATE": return 3000000;
    case "ALT_COST_SURCHARGE": return 0.23;
    default: return null;
  }
}

function getItemEstimateValue(item, code, fallback = null) {
  const itemValues = state.itemEstimateValues[item.key] || {};
  const saved = itemValues[code];
  const targetedField = (TARGETED_PLAN_INPUTS[item.id] || [])
    .find((candidate) => candidate.code === code);
  if (targetedField && Object.hasOwn(itemValues, code)) return saved;
  if (code === JKSO_CONSTRUCTION_INPUT.code) {
    const value = String(saved ?? JKSO_CONSTRUCTION_INPUT.options[0]);
    const codeMatch = value.match(/^\s*(\d+)/)?.[1];
    const availableOptions = getRtsConstructionOptions(item);
    return availableOptions.includes(codeMatch)
      ? codeMatch : availableOptions[0] ?? JKSO_CONSTRUCTION_INPUT.options[0];
  }
  if (saved !== null && saved !== undefined) return saved;
  if (code === "BOILER_UNIT_COST") {
    const moduleCapacity = numericOrNull(getItemEstimateValue(item, "BOILER_UNIT_CAPACITY_KW"));
    return moduleCapacity === null ? null : moduleCapacity * 10000;
  }
  if (code === HALL_COVERAGE_INPUT.code) {
    return state.secondaryValues["Veřejná krytá hala – způsob pokrytí"]
      || state.defaults.secondary[HALL_COVERAGE_INPUT.code]
      || HALL_COVERAGE_INPUT.options[0];
  }
  const alternativeField = ALTERNATIVE_ESTIMATE_INPUTS.find((candidate) => candidate.code === code);
  if (alternativeField) return getAlternativeEstimateDefault(code);
  if (targetedField) return targetedField.defaultValue ?? fallback;
  const field = ITEM_ESTIMATE_INPUTS.find((candidate) => candidate.code === code);
  return field ? getItemEstimateDefault(item, code) : fallback;
}

function getRtsBuildingGroup(item) {
  return RTS_801_GROUP_BY_ITEM_ID[item.id];
}

function getRtsConstructionOptions(item) {
  const rates = RTS_801_CONSTRUCTION_RATES[getRtsBuildingGroup(item)];
  return rates ? Object.keys(rates).sort((left, right) => Number(left) - Number(right)) : [];
}

function getRtsConstructionRate(item) {
  const buildingGroup = getRtsBuildingGroup(item);
  if (!buildingGroup) return undefined;
  const constructionCode = getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code);
  const rates = RTS_801_CONSTRUCTION_RATES[buildingGroup];
  return Object.prototype.hasOwnProperty.call(rates, constructionCode)
    ? rates[constructionCode] : null;
}

function getTargetedPlanCostRate(item) {
  return getRtsBuildingGroup(item)
    ? numericOrNull(getRtsConstructionRate(item))
    : getTargetedPlanInput(item, "MODEL_COST_PER_M3");
}

function getProjectEstimateValue(item, code) {
  const saved = state.itemEstimateValues[item.key]?.[code];
  if (code === "POLY_CENA_M3" && getRtsBuildingGroup(item)) {
    const constructionRate = getRtsConstructionRate(item);
    return numericOrNull(constructionRate);
  }
  if (saved !== null && saved !== undefined) return numericOrNull(saved);
  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id) && code === "K_PLOCHA_AREALU") {
    const usableArea = numericOrNull(item.usableArea);
    const parcelArea = numericOrNull(item.parcelArea);
    const totalFloorFactor = getProjectEstimateValue(item, "K_CELKOVA_PLOCHA");
    return usableArea > 0 && parcelArea > 0 && totalFloorFactor > 0
      ? parcelArea / (usableArea * totalFloorFactor) : numericOrNull(getSecondary(code));
  }
  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id) && code === "K_ZPEVNENI") {
    const usableArea = numericOrNull(item.usableArea);
    const pavedArea = numericOrNull(item.pavedArea);
    const totalFloorFactor = getProjectEstimateValue(item, "K_CELKOVA_PLOCHA");
    const siteAreaFactor = getProjectEstimateValue(item, "K_PLOCHA_AREALU");
    const footprintFactor = getProjectEstimateValue(item, "K_ZASTAVENI");
    const siteArea = usableArea * totalFloorFactor * siteAreaFactor;
    const footprintArea = usableArea * totalFloorFactor * footprintFactor;
    const availableArea = siteArea - footprintArea;
    return pavedArea > 0 && availableArea > 0
      ? pavedArea / availableArea : numericOrNull(getSecondary(code));
  }
  return numericOrNull(getSecondary(code));
}

function getTargetedPlanInput(item, code) {
  const saved = state.itemEstimateValues[item.key]?.[code];
  if (saved === "" || saved === null) return null;
  if (saved === undefined) {
    const field = (TARGETED_PLAN_INPUTS[item.id] || [])
      .find((candidate) => candidate.code === code);
    return numericOrNull(field?.defaultValue);
  }
  return numericOrNull(saved);
}

function getTargetedPlanSource(item, code) {
  return String(state.itemEstimateValues[item.key]?.[code] ?? "").trim();
}

function getTargetedPlanContext() {
  const derived = derivedValues();
  const areaCode = getSecondary("Plocha pro odhad dostupnosti") || "S_zastav";
  const areaHa = numericOrNull(getPrimary(areaCode));
  const projectPopulation = numericOrNull(derived.O_uzemi_cil);
  const catchmentPopulation = numericOrNull(derived.O_prah) ?? projectPopulation;
  return {
    projectPopulation,
    catchmentPopulation,
    territory: getResidentParkingTerritoryType(),
    areaCode,
    areaHa,
    populationDensity: projectPopulation !== null && areaHa > 0
      ? projectPopulation / areaHa : null
  };
}

function getTargetedCoverageEstimate(distanceMeters) {
  const context = getTargetedPlanContext();
  const shapeCode = getPrimary("KOD_TVAR");
  const shape = GEOMETRY_EXTENTS[shapeCode];
  const overlap = numericOrNull(getSecondary(COVERAGE_OVERLAP_INPUT.code)) ?? 0.25;
  const detour = numericOrNull(state.pedestrianDetourFactor);
  if (!(context.areaHa > 0) || !shape || !(detour > 0)
    || !(distanceMeters > 0) || overlap < 0 || overlap >= 1) return null;
  const boundingArea = context.areaHa * 10000 * shape.areaFactor;
  const longSide = Math.sqrt(boundingArea * shape.aspect);
  const shortSide = Math.sqrt(boundingArea / shape.aspect);
  const radius = distanceMeters / detour;
  const spacing = Math.SQRT2 * radius * (1 - overlap);
  const columns = Math.max(1, Math.ceil(longSide / spacing));
  const rows = Math.max(1, Math.ceil(shortSide / spacing));
  return { ...context, shapeCode, distanceMeters, overlap, detour, columns, rows, count: columns * rows };
}

function getTargetedFacilityPlan(item, quantity) {
  const assumptions = TARGETED_FACILITY_ASSUMPTIONS[item.id];
  if (!assumptions || quantity === null || quantity === undefined) return null;
  const modelFields = (TARGETED_PLAN_INPUTS[item.id] || [])
    .filter((field) => field.code.startsWith("MODEL_"));
  if (modelFields.some((field) => getTargetedPlanInput(item, field.code) === null)) return null;
  const input = (code) => getTargetedPlanInput(item, code);

  let capacity = Math.max(0, quantity);
  let units = capacity;
  let unitLabel = assumptions.unitLabel;
  let usableArea = 0;
  let totalFloorArea = 0;
  let siteArea = 0;
  let pavedArea = 0;
  let detail;
  const areaPerUnit = input("MODEL_AREA_PER_UNIT");
  const floorFactor = input("MODEL_FLOOR_FACTOR");
  const sitePerUnit = input("MODEL_SITE_PER_UNIT");
  const pavedShare = input("MODEL_PAVED_SHARE");
  let volumePerFloorArea = input("MODEL_VOLUME_PER_FLOOR_M2");
  const rate = getTargetedPlanCostRate(item);
  const surcharge = input("MODEL_COST_SURCHARGE");

  if (item.id === "VYB_09D") {
    const vehicles = capacity;
    const vehiclesPerBase = input("MODEL_VEHICLES_PER_BASE");
    const bases = Math.ceil(vehicles / vehiclesPerBase);
    units = bases;
    usableArea = bases * input("MODEL_BASE_AREA") + vehicles * areaPerUnit;
    totalFloorArea = usableArea;
    siteArea = bases * sitePerUnit;
    pavedArea = bases * input("MODEL_PAVED_PER_BASE") + vehicles * input("MODEL_PAVED_PER_UNIT");
    unitLabel = "výjezdových základen";
    capacity = vehicles;
    detail = `${formatNumber(vehicles)} výjezdových vozů; ${formatNumber(vehiclesPerBase)} vozů/základnu`;
  } else if (item.id === "VYB_16A") {
    const vehicles = capacity;
    const vehiclesPerBase = input("MODEL_VEHICLES_PER_BASE");
    const stations = Math.ceil(vehicles / vehiclesPerBase);
    units = stations;
    usableArea = stations * input("MODEL_BASE_AREA") + vehicles * areaPerUnit;
    totalFloorArea = usableArea;
    siteArea = stations * sitePerUnit;
    pavedArea = stations * input("MODEL_PAVED_PER_BASE") + vehicles * input("MODEL_PAVED_PER_UNIT");
    unitLabel = "hasičských stanic";
    capacity = vehicles;
    detail = `${formatNumber(vehicles)} požárních vozidel; ${formatNumber(vehiclesPerBase)} vozidel/stanici`;
  } else if (item.id === "VYB_17A") {
    const officers = capacity;
    const officersPerStation = input("MODEL_CAPACITY_PER_FACILITY");
    const stations = Math.ceil(officers / officersPerStation);
    units = stations;
    usableArea = Math.max(stations * input("MODEL_BASE_AREA"), officers * areaPerUnit);
    totalFloorArea = usableArea;
    siteArea = stations * sitePerUnit;
    pavedArea = stations * input("MODEL_PAVED_PER_BASE");
    unitLabel = "služeben";
    detail = `${formatNumber(officers)} policistů; nejvýše ${formatNumber(officersPerStation)} policistů/služebnu`;
  } else if (item.id === "VYB_17B") {
    const people = capacity;
    const constructionHeight = input("MODEL_CONSTRUCTION_HEIGHT");
    usableArea = people * areaPerUnit;
    totalFloorArea = usableArea * floorFactor;
    siteArea = people * sitePerUnit;
    pavedArea = siteArea * pavedShare;
    unitLabel = "osob";
    detail = `${formatNumber(people)} osob × ${formatNumber(areaPerUnit)} m²/osobu`;
    volumePerFloorArea = constructionHeight;
  } else {
    usableArea = capacity * areaPerUnit;
    totalFloorArea = usableArea * floorFactor;
    siteArea = capacity * sitePerUnit;
    pavedArea = siteArea * pavedShare;
    detail = `${formatNumber(capacity)} ${unitLabel} × ${formatNumber(areaPerUnit)} m²/${unitLabel}`;
  }

  const volume = totalFloorArea * volumePerFloorArea;
  const footprintArea = totalFloorArea / (floorFactor > 0 ? floorFactor : 1);
  const volumeFactorUnit = item.id === "VYB_17B" ? "m" : "m³/m²";
  const baseCost = volume * rate;
  const cost = Math.round(baseCost * (1 + surcharge));
  const count = ["VYB_09D", "VYB_16A", "VYB_17A"].includes(item.id)
    ? units : Math.ceil(capacity / input(item.id === "VYB_01_J"
      ? "MODEL_GROUP_CAPACITY" : "MODEL_CAPACITY_PER_FACILITY"));
  const countForms = {
    VYB_01_J: ["skupina", "skupiny", "skupin"],
    VYB_02C_1: ["škola", "školy", "škol"],
    VYB_02C_2: ["škola", "školy", "škol"],
    VYB_09A_4: ["ordinace", "ordinace", "ordinací"],
    VYB_09D: ["základna", "základny", "základen"],
    VYB_16A: ["stanice", "stanice", "stanic"],
    VYB_17A: ["služebna", "služebny", "služeben"],
    VYB_17B: ["úkryt", "úkryty", "úkrytů"]
  }[item.id];
  const capacityForms = {
    VYB_01_J: ["místo", "místa", "míst"],
    VYB_02C_1: ["žákovské místo", "žákovská místa", "žákovských míst"],
    VYB_02C_2: ["žákovské místo", "žákovská místa", "žákovských míst"],
    VYB_09D: ["výjezdový vůz", "výjezdové vozy", "výjezdových vozů"],
    VYB_16A: ["požární vozidlo", "požární vozidla", "požárních vozidel"],
    VYB_17A: ["policista", "policisté", "policistů"],
    VYB_17B: ["osoba", "osoby", "osob"]
  }[item.id];
  const capacityDisplay = item.id === "VYB_09A_4"
    ? `${formatNumber(capacity)} souč. zátěže`
    : `${formatNumber(capacity)} ${czechCountUnit(capacity, ...capacityForms)}`;
  const countDisplay = `${formatNumber(count)} ${czechCountUnit(count, ...countForms)}`;
  const countMetricLabel = {
    VYB_01_J: "Modelový počet dětských skupin",
    VYB_02C_1: "Modelový počet školních zařízení",
    VYB_02C_2: "Modelový počet ZUŠ",
    VYB_09A_4: "Modelový počet ordinačních modulů",
    VYB_09D: "Modelový počet výjezdových základen",
    VYB_16A: "Modelový počet hasičských stanic",
    VYB_17A: "Modelový počet policejních služeben",
    VYB_17B: "Modelový počet úkrytových modulů"
  }[item.id];
  const volumeLabel = item.id === "VYB_17B"
    ? "Orientační obestavěný prostor při zadané konstrukční výšce"
    : "Orientační obestavěný prostor";

  return {
    metrics: [
      { label: "Návrhová kapacita", value: capacityDisplay, unit: "" },
      { label: countMetricLabel, value: countDisplay, unit: "" },
      { label: "Orientační užitná plocha", value: usableArea, unit: "m²" },
      { label: "Orientační celková podlahová plocha", value: totalFloorArea, unit: "m²" },
      { label: "Orientační zastavěná plocha", value: footprintArea, unit: "m²" },
      { label: volumeLabel, value: volume, unit: "m³" },
      { label: "Orientační plocha areálu", value: siteArea, unit: "m²" },
      { label: "Orientační zpevněná plocha", value: pavedArea, unit: "m²" },
      { label: "Orientační investice včetně VRN a rezervy", value: cost, unit: "Kč" }
    ],
    equations: [
      `Návrhová plocha = ${detail} = ${formatNumber(usableArea)} m² užitné plochy; obestavěný prostor = ${formatNumber(totalFloorArea)} m² × ${formatNumber(volumePerFloorArea)} ${volumeFactorUnit} = ${formatNumber(volume)} m³.`,
      `Orientační investice = ${formatNumber(volume)} m³ × ${formatNumber(rate)} Kč/m³ × (1 + ${formatNumber(surcharge * 100)} % VRN a rezerva) = ${formatNumber(cost)} Kč.`
    ],
    footprintArea,
    pavedArea,
    source: `${assumptions.rateSource}${getRtsBuildingGroup(item)
      ? `; ${RTS_801_SOURCE} JKSO ${getRtsBuildingGroup(item)}, konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}, sazba ${formatNumber(rate)} Kč/m³ OP`
      : ""}. Plošné, kapacitní a cenové koeficienty jsou editovatelné scénářové předpoklady; nezahrnují cenu pozemku ani individuální technologie.`
  };
}

function addTargetedCoverageMetrics(metrics, distanceMeters, label = "Orientační počet bodů pokrytí") {
  const coverage = getTargetedCoverageEstimate(distanceMeters);
  const context = getTargetedPlanContext();
  metrics.push(
    { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
    { label: "Orientační hustota obyvatel", value: context.populationDensity, unit: "obyvatel/ha" },
    { label, value: coverage?.count ?? null, unit: "bodů" }
  );
  return coverage;
}

function getTargetedPlanEstimate(item) {
  const context = getTargetedPlanContext();
  const { projectPopulation, catchmentPopulation, territory } = context;
  const estimate = (metrics, equations, sources, facilityNeed = null, modelEquations = null) => {
    const facilityPlan = getTargetedFacilityPlan(item, facilityNeed);
    if (facilityPlan) {
      metrics.push(...facilityPlan.metrics);
      sources.push(facilityPlan.source);
    }
    const calculationEquations = modelEquations
      ? [...modelEquations, ...(facilityPlan?.equations || [])]
      : [...equations, ...(facilityPlan?.equations || [])];
    const verificationEquations = modelEquations
      ? equations.filter((equation) => !modelEquations.includes(equation)) : [];
    return {
      type: "targeted", metrics, equations: calculationEquations,
      verificationEquations, sources, facilityPlan
    };
  };

  if (item.id === "VYB_01_J") {
    const demographicShare = getTargetedPlanInput(item, "MODEL_DEMOGRAPHIC_SHARE");
    const attendanceShare = getTargetedPlanInput(item, "MODEL_ATTENDANCE_SHARE");
    const groupSize = getTargetedPlanInput(item, "MODEL_GROUP_CAPACITY");
    const minimumGroupSize = getTargetedPlanInput(item, "MODEL_MIN_GROUP_CAPACITY");
    const children = catchmentPopulation === null || demographicShare === null
      ? null : Math.ceil(catchmentPopulation * demographicShare);
    const expectedAttendance = children === null || attendanceShare === null
      ? null : Math.ceil(children * attendanceShare);
    const existingCapacity = getTargetedPlanInput(item, "EARLY_CHILDCARE_EXISTING_CAPACITY");
    const existingDistance = getTargetedPlanInput(item, "EARLY_CHILDCARE_ACCESS_DISTANCE");
    const source = getTargetedPlanSource(item, "EARLY_CHILDCARE_DATA_SOURCE");
    const coverageDistance = territory === "D" ? 600 : 400;
    const verifiedCapacity = existingCapacity !== null && Number.isInteger(existingCapacity)
      && existingDistance !== null && existingDistance <= coverageDistance && source.length > 0;
    const unmet = expectedAttendance === null ? null
      : verifiedCapacity ? Math.max(0, expectedAttendance - existingCapacity) : expectedAttendance;
    const groups = unmet === null || unmet === 0 ? 0
      : groupSize === null ? null : Math.ceil(unmet / groupSize);
    const groupCapacity = groups === null || minimumGroupSize === null ? null
      : unmet === 0 ? 0
        : Math.max(unmet, groups * Math.min(groupSize, minimumGroupSize));
    const groupUnit = groups === null ? "skupin"
      : czechCountUnit(groups, "skupina", "skupiny", "skupin");
    const metrics = [
      { label: `Odhad dětí ve věku 0–3 roky (${demographicShare === null ? "nevyplněno" : `${formatNumber(demographicShare * 100)} %`} populace)`, value: children, unit: "dětí" },
      { label: `Odhad poptávky po péči (${attendanceShare === null ? "nevyplněno" : `${formatNumber(attendanceShare * 100)} %`} věkové kohorty)`, value: expectedAttendance, unit: "míst" },
      { label: "Obyvatelé řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Obyvatelé včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Započtená ověřená volná kapacita okolí", value: verifiedCapacity ? existingCapacity : null, unit: "míst" },
      { label: "Skutečná pěší vzdálenost k této kapacitě", value: existingDistance, unit: "m" },
      { label: "Poptávka nepokrytá ověřenou kapacitou", value: unmet, unit: "míst" },
      { label: `Navrhovaný počet dětských skupin (${formatNumber(groupSize)} míst/skupinu)`, value: unmet === null ? null : groups, unit: groupUnit },
      { label: "Návrhová kapacita skupin", value: unmet === null ? null : groupCapacity, unit: "míst" }
    ];
    const coverage = addTargetedCoverageMetrics(metrics, coverageDistance);
    return estimate(metrics, [
      children === null
        ? "Chybí počet obyvatel; odhad cílové věkové kohorty nelze vypočítat."
        : `Věková kohorta 0–3 roky (scénář) = ${formatNumber(catchmentPopulation)} obyvatel × ${formatNumber(demographicShare * 100)} % = ${formatNumber(children)} dětí; odhad využívajících péči = ${formatNumber(children)} × ${formatNumber(attendanceShare * 100)} % = ${formatNumber(expectedAttendance)} míst.`,
      verifiedCapacity && unmet !== null
        ? `Návrhová nepokrytá poptávka = max(0, ${formatNumber(expectedAttendance)} − ${formatNumber(existingCapacity)} ověřených míst) = ${formatNumber(unmet)} míst; dostupná kapacita je v dosahu ${coverageDistance} m.`
        : existingCapacity !== null && existingDistance !== null && existingDistance > coverageDistance
          ? `Zadaná kapacita ${formatNumber(existingCapacity)} míst se nezapočítá: zařízení je ve skutečné pěší vzdálenosti ${formatNumber(existingDistance)} m, mimo dosah ${coverageDistance} m.`
          : `Bez ověření okolní nabídky se pro návrh použije celá modelová poptávka ${formatNumber(unmet)} míst; jde o předběžnou potřebu, nikoli potvrzený deficit.`,
      `Počet skupin = ceil(${formatNumber(unmet)} míst ÷ ${formatNumber(groupSize)} míst/skupinu) = ${formatNumber(groups)}; kapacita = max(${formatNumber(unmet)} míst, ${formatNumber(groups)} × minimálně ${formatNumber(minimumGroupSize)} míst/skupinu) = ${formatNumber(groupCapacity)} míst.`,
      coverage
        ? `Plošná kontrola dosahu: ${formatNumber(context.areaHa)} ha, typ území ${territory}, pěší dosah ${coverageDistance} m; orientační síť ${coverage.columns} × ${coverage.rows} = ${coverage.count} bodů. Jde o geometrický odhad, ne o povinný počet zařízení.`
        : `Plošnou dostupnost nelze odhadnout bez plochy, tvaru a koeficientu pěší sítě; pro typ ${territory} se používá orientační dosah ${coverageDistance} m.`,
      `${demographicShare === null ? "Podíl věkové kohorty není vyplněn." : `Podíl kohorty ${formatNumber(demographicShare * 100)} %`}${attendanceShare === null ? " Podíl účasti není vyplněn." : ` a účast ${formatNumber(attendanceShare * 100)} %`} Jde o editovatelné scénářové předpoklady; nahraďte je místními demografickými a provozními daty.`
    ], [
      "Pracovní sešit Cluster_Skolstvi uvádí pěší dosah 300–400 m v typech A–C, 400–600 m v typu D, dětské skupiny 12–24 míst a 10–15 m²/dítě. Aplikace jako návrhovou hodnotu používá 12,5 m²/místo a 24 míst/skupinu; věkový podíl 4,2 % a využití 30 % jsou scénářové odhady, nikoli normy. Plošný model nezohledňuje skutečné trasy ani existující polohy.",
      ...(source ? [`Zadaný zdroj vstupů: ${source}.`] : [])
    ], unmet === null ? null : groupCapacity, [children === null || expectedAttendance === null
      ? "Chybí počet obyvatel nebo demografický podíl; poptávku nelze vypočítat."
      : `Poptávka = ceil(${formatNumber(catchmentPopulation)} obyvatel × ${formatNumber(demographicShare * 100)} % × ${formatNumber(attendanceShare * 100)} %) = ${formatNumber(expectedAttendance)} míst.`]);
  }

  if (item.id === "VYB_02C_1" || item.id === "VYB_02C_2") {
    const secondarySchool = item.id === "VYB_02C_1";
    const freeSeatsCode = secondarySchool ? "SS_EXISTING_FREE_SEATS" : "ZUS_EXISTING_FREE_SEATS";
    const sourceCode = secondarySchool ? "SS_CAPACITY_SOURCE" : "ZUS_CAPACITY_SOURCE";
    const studentShare = getTargetedPlanInput(item, "MODEL_DEMOGRAPHIC_SHARE");
    const freeSeats = getTargetedPlanInput(item, freeSeatsCode);
    const seatsPerSchool = getTargetedPlanInput(item, "MODEL_CAPACITY_PER_FACILITY");
    const source = getTargetedPlanSource(item, sourceCode);
    const schoolVhdMinutes = secondarySchool
      ? getTargetedPlanInput(item, "SS_VHD_MINUTES") : null;
    const artWalkDistance = secondarySchool
      ? null : getTargetedPlanInput(item, "ZUS_WALK_DISTANCE");
    const artVhdMinutes = secondarySchool
      ? null : getTargetedPlanInput(item, "ZUS_VHD_MINUTES");
    const studentDemand = catchmentPopulation !== null && studentShare !== null
      ? Math.ceil(catchmentPopulation * studentShare) : null;
    const accessVerified = Boolean(source) && (
      secondarySchool ? schoolVhdMinutes !== null && ["A", "B"].includes(territory)
        : territory === "D" ? artVhdMinutes !== null : artWalkDistance !== null);
    const accessPasses = accessVerified && (
      secondarySchool ? schoolVhdMinutes <= 45
        : territory === "D" ? artVhdMinutes <= 30 : artWalkDistance <= 1000);
    const capacityVerified = Boolean(source) && studentDemand !== null
      && freeSeats !== null && Number.isInteger(freeSeats) && freeSeats >= 0
      && accessVerified && accessPasses;
    const unmetSeats = studentDemand === null ? null : capacityVerified
      ? Math.max(0, studentDemand - freeSeats) : studentDemand;
    const newSchools = unmetSeats === null ? null : Math.ceil(unmetSeats / seatsPerSchool);
    const metrics = [
      { label: "Počet obyvatel řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Počet obyvatel včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Typ území", value: territory, unit: "" }
    ];
    const coverage = secondarySchool || territory === "D"
      ? null : addTargetedCoverageMetrics(metrics, 1000);
    if (secondarySchool || territory === "D") metrics.push(
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" }
    );
    metrics.push(
      { label: secondarySchool
        ? `Odhad cílové středoškolské kohorty (${studentShare === null ? "nevyplněno" : `${formatNumber(studentShare * 100)} %`} populace)`
        : `Odhad využívajících ZUŠ (${studentShare === null ? "nevyplněno" : `${formatNumber(studentShare * 100)} %`} populace)`,
        value: studentShare === null ? null : studentShare * 100, unit: "%" },
      { label: "Odhadovaná poptávka po místech", value: studentDemand, unit: "žáků" },
      { label: "Ověřená volná kapacita", value: capacityVerified ? freeSeats : null, unit: "míst" },
      { label: "Poptávka po odečtu ověřené volné kapacity", value: unmetSeats, unit: "míst" },
      { label: "Počet modelových zařízení pro nepokrytou poptávku", value: newSchools, unit: "škol" },
      { label: "Ověřená dostupnost", value: accessVerified
        ? accessPasses ? "Vyhovuje" : "Nevyhovuje"
        : "Neověřeno", unit: "" },
      ...(secondarySchool
        ? [{ label: "Ověřený čas VHD k SŠ", value: schoolVhdMinutes, unit: "min" }]
        : [
          { label: "Ověřená pěší vzdálenost k ZUŠ", value: artWalkDistance, unit: "m" },
          { label: "Ověřený čas VHD v mikroregionu", value: artVhdMinutes, unit: "min" }
        ])
    );
    return estimate(metrics, [
      studentDemand === null
        ? "Chybí počet obyvatel nebo demografický podíl; poptávku po místech nelze vypočítat."
        : `Poptávka po místech = ceil(${formatNumber(catchmentPopulation)} obyvatel × ${formatNumber(studentShare * 100)} %) = ${formatNumber(studentDemand)} žáků.`,
      secondarySchool
        ? accessVerified
          ? `Skutečný čas VHD k SŠ = ${formatNumber(schoolVhdMinutes)} min; limit pro typ A/B je 45 min, výsledek ${accessPasses ? "vyhovuje" : "nevyhovuje"}.`
          : `Dostupnost SŠ se v typech A/B ověřuje skutečnou cestou VHD do 45 min; projektová plocha ${context.areaHa === null ? "chybí" : `${formatNumber(context.areaHa)} ha`} se nepřevádí na dobu jízdy.`
        : coverage
          ? `ZUŠ: orientační síť pro 1 000 m pěší dosah nad ${formatNumber(context.areaHa)} ha je ${coverage.columns} × ${coverage.rows} = ${coverage.count} bodů; ověřit skutečnou trasu a existující školu.`
          : territory === "D"
            ? "U typu D se dostupnost ZUŠ ověřuje v mikroregionu do 20–30 minut VHD; plocha se nepřevádí na dobu jízdy."
          : `ZUŠ: plošné pokrytí 1 000 m nelze odhadnout bez plochy, tvaru a koeficientu pěší sítě.`,
      !secondarySchool
        ? territory === "D"
          ? accessVerified
            ? `Skutečný čas VHD k ZUŠ = ${formatNumber(artVhdMinutes)} min; pracovní limit mikroregionu je nejvýše 30 min, výsledek ${accessPasses ? "vyhovuje" : "nevyhovuje"}.`
            : "Pro typ D ověřte skutečnou dostupnost ZUŠ v mikroregionu do 20–30 minut VHD."
          : accessVerified
            ? `Skutečná pěší vzdálenost k ZUŠ = ${formatNumber(artWalkDistance)} m; limit je 1 000 m, výsledek ${accessPasses ? "vyhovuje" : "nevyhovuje"}.`
            : "Skutečná pěší trasa k ZUŠ nebyla doložena; geometrická síť ji nenahrazuje."
        : "",
      capacityVerified
        ? `Nepokrytá poptávka = max(0, ceil(${formatNumber(studentDemand)} žáků) − ${formatNumber(freeSeats)} volných míst) = ${unmetSeats} míst; při ${formatNumber(seatsPerSchool)} místech/školu jde o ${newSchools} zařízení k prověření.`
        : `Modelová nepokrytá poptávka = ${formatNumber(studentDemand)} míst; bez ověřené dostupné volné kapacity se pro návrh neuplatní žádný odečet. Z toho vychází ${newSchools} modelových zařízení po ${seatsPerSchool} míst.`,
      `${studentShare === null ? "Poptávkový podíl není vyplněn." : `Poptávkový podíl ${formatNumber(studentShare * 100)} %`}${seatsPerSchool === null ? " Kapacita zařízení není vyplněna." : ` a kapacita ${formatNumber(seatsPerSchool)} míst na zařízení`} jsou editovatelné scénáře, nikoli právní kvóty.`
    ], [
      `Metodika MMR/TB050MMR01 stanoví dostupnost SŠ v typech A/B do 45 minut VHD; pracovní sešit uvádí u ZUŠ docházku do 1 000 m a dostupnost v mikroregionu. Poptávkový podíl a kapacita jsou editovatelné odhady, nikoli právní normy.${source ? ` Zadaný zdroj: ${source}.` : ""}`
    ], unmetSeats === null ? null : unmetSeats, [studentDemand === null
      ? "Chybí počet obyvatel nebo demografický podíl; poptávku nelze vypočítat."
      : `Poptávka po místech = ceil(${formatNumber(catchmentPopulation)} obyvatel × ${formatNumber(studentShare * 100)} %) = ${formatNumber(studentDemand)} žáků.`]);
  }

  if (item.id === "VYB_09A_4") {
    const patientsPerFte = getTargetedPlanInput(item, "MODEL_POPULATION_PER_FTE");
    const population = catchmentPopulation;
    const fte = population !== null && patientsPerFte ? population / patientsPerFte : null;
    const twoDayThreshold = patientsPerFte ? patientsPerFte * 0.4 : null;
    const availableFte = getTargetedPlanInput(item, "GYNECOLOGY_AVAILABLE_FTE");
    const walkDistance = getTargetedPlanInput(item, "GYNECOLOGY_WALK_DISTANCE");
    const driveMinutes = getTargetedPlanInput(item, "GYNECOLOGY_DRIVE_MINUTES");
    const source = getTargetedPlanSource(item, "GYNECOLOGY_ACCESS_SOURCE");
    const physicalAccess = ["A", "B", "C"].includes(territory);
    const coverage = physicalAccess ? getTargetedCoverageEstimate(600) : null;
    const municipalityThresholdMet = physicalAccess
      && projectPopulation !== null && projectPopulation > 2000;
    const accessVerified = Boolean(source) && (
      physicalAccess ? walkDistance !== null : driveMinutes !== null);
    const accessPasses = accessVerified && (
      physicalAccess ? walkDistance <= 600 : driveMinutes <= 35);
    const capacityVerified = Boolean(source) && availableFte !== null
      && accessVerified && accessPasses;
    const remainingFte = fte === null ? null : capacityVerified
      ? Math.max(0, fte - availableFte) : fte;
    return estimate([
      { label: "Počet obyvatel včetně spádu po záměru", value: population, unit: "obyvatel" },
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Populační práh přítomnosti v obci (> 2 000)", value: physicalAccess
        ? municipalityThresholdMet ? "Dosažen (ověřit populaci obce)" : "Nedosažen / ověřit"
        : "Mimo strukturální standard A/B/C", unit: "" },
      { label: "Indikativní souč. zátěž ordinací", value: fte, unit: "souč. zátěže" },
      { label: "Orientační práh pro dvoudenní ordinaci (40 % poměru)", value: twoDayThreshold, unit: "obyvatel" },
      { label: "Ověřená volná kapacita", value: capacityVerified ? availableFte : null, unit: "souč. zátěže" },
      { label: "Poptávka po gynekologii k ověření / po odečtu kapacity", value: remainingFte, unit: "souč. zátěže" },
      { label: "Kontrola dostupnosti", value: accessVerified
        ? accessPasses ? "Vyhovuje" : "Nevyhovuje"
        : "Neověřeno", unit: "" },
      ...(physicalAccess
        ? [
          { label: "Ověřená pěší vzdálenost", value: walkDistance, unit: "m" },
          { label: "Orientační body sítě při pěším dosahu 600 m", value: coverage?.count ?? null, unit: "bodů" }
        ]
        : [
          { label: "Ověřený dojezd (typ D)", value: driveMinutes, unit: "min" },
          { label: "Standard dostupnosti pro typ D", value: "35 min dojezdu", unit: "" }
        ])
    ], [
      population === null ? "Chybí počet obyvatel ve spádu; personální indikátor nelze vypočítat."
        : `Indikativní potřeba = ${formatNumber(population)} obyvatel ve spádu ÷ ${formatNumber(patientsPerFte)} obyvatel na jednotku souč. zátěže = ${formatNumber(fte)} jednotek souč. zátěže.`,
      `Orientační populační práh pro dvoudenní ordinační dobu = 40 % × ${formatNumber(patientsPerFte)} = ${formatNumber(twoDayThreshold)} obyvatel.`,
      physicalAccess
        ? `Populační práh MMR pro přítomnost v obci (> 2 000 obyvatel) ${municipalityThresholdMet ? "je dosažen podle zadané populace řešeného území" : "není dosažen podle zadané populace řešeného území"}; ověřte, že tento vstup odpovídá celé obci.`
        : "Pro typ D se místo strukturálního populačního prahu uplatňuje dojezdový standard.",
      capacityVerified && remainingFte !== null
        ? `Poptávka po odečtu ověřené volné kapacity dostupné ordinace = max(0, ${formatNumber(fte)} − ${formatNumber(availableFte)}) = ${formatNumber(remainingFte)} jednotek souč. zátěže.`
        : `Poptávka k vyhledání v okolí = ${formatNumber(fte)} jednotek souč. zátěže; bez ověřené kapacity dostupné ordinace se nic neodečítá.`,
      physicalAccess
        ? coverage
          ? `Plošný odhad pokrytí: ${formatNumber(context.areaHa)} ha, typ ${territory}, 600 m pěšky; mřížka ${coverage.columns} × ${coverage.rows} = ${coverage.count} bodů.`
          : "Plošný odhad 600 m nelze vypočítat bez plochy, tvaru a koeficientu pěší sítě."
        : "Pro typ D stanoví metodika časovou dostupnost dojezdem 35 minut; nelze ji převést z geometrie plochy bez síťové analýzy.",
      accessVerified
        ? physicalAccess
          ? `Skutečná pěší vzdálenost ${formatNumber(walkDistance)} m ${accessPasses ? "splňuje" : "překračuje"} limit 600 m.`
          : `Ověřený dojezd ${formatNumber(driveMinutes)} min ${accessPasses ? "splňuje" : "překračuje"} limit 35 min.`
        : "Dostupnost poskytovatele musí být ověřena podle typu území a doložena zdrojem.",
      `Souč. zátěž se převádí na výměru ordinací podle editovatelného scénáře ${formatNumber(getTargetedPlanInput(item, "MODEL_AREA_PER_UNIT"))} m²/jednotku souč. zátěže; výměra není právním normativem.`
    ], [
      "Metodika MMR/TB050MMR01, s. 9, stanoví dostupnost gynekologické péče v typech A/B/C do 600 m a v typu D do 35 minut. Historický indikátor 14 000 obyvatel na jednotku souč. zátěže je uveden v TB050MMR01, s. 26–27 a pracovním listu Cluster_Zdravotnictvi; není to současná kvóta ani doporučení ČSÚ. Zadaný modelový poměr je editovatelný."
      ,
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], remainingFte, [population === null || !patientsPerFte
      ? "Chybí počet obyvatel nebo počet pacientek na jednotku souč. zátěže; potřebu nelze vypočítat."
      : `Potřeba souč. zátěže = ${formatNumber(population)} obyvatel ve spádu ÷ ${formatNumber(patientsPerFte)} obyvatel/jednotku souč. zátěže = ${formatNumber(fte)} jednotek souč. zátěže.`]);
  }

  if (item.id === "VYB_09D") {
    const covered = getTargetedPlanInput(item, "ZZS_COVERED_POPULATION");
    const responseMinutes = getTargetedPlanInput(item, "ZZS_MAX_RESPONSE_MINUTES");
    const source = getTargetedPlanSource(item, "ZZS_COVERAGE_SOURCE");
    const populationPerVehicle = getTargetedPlanInput(item, "MODEL_POPULATION_PER_VEHICLE");
    const vehiclesPerBase = getTargetedPlanInput(item, "MODEL_VEHICLES_PER_BASE");
    const validResponse = Boolean(source) && responseMinutes !== null && catchmentPopulation !== null;
    const coveredIsValid = covered !== null && catchmentPopulation !== null && covered <= catchmentPopulation;
    const uncovered = coveredIsValid ? Math.max(0, catchmentPopulation - covered) : catchmentPopulation;
    const needsBaseReview = !validResponse || responseMinutes > 20;
    const responseVehicles = catchmentPopulation === null || !needsBaseReview
      ? 0 : Math.max(vehiclesPerBase, Math.ceil(Math.max(0, uncovered) / populationPerVehicle));
    return estimate([
      { label: "Obyvatelé řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Obyvatelé včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Ověřená nejdelší dojezdová doba v projektovém spádu", value: validResponse ? responseMinutes : null, unit: "min" },
      { label: "Ověřeně pokrytí obyvatelé", value: coveredIsValid ? covered : null, unit: "obyvatel" },
      { label: "Obyvatelé bez doloženého pokrytí", value: uncovered, unit: "obyvatel" },
      { label: `Modelová poptávka po výjezdových vozech (1 / ${formatNumber(populationPerVehicle)} obyvatel)`,
        value: `${formatNumber(needsBaseReview ? responseVehicles : 0)} ${czechCountUnit(needsBaseReview ? responseVehicles : 0, "vůz", "vozy", "vozů")}`, unit: "" },
      { label: "Výsledek podmíněného posouzení", value: needsBaseReview ? "Potřebu základny započítat k ověření" : "Nová základna se nezapočítá", unit: "" }
    ], [
      `Požadovaná mezní dojezdová doba pro projektový spád = 20 min; ${validResponse ? `ověřeno ${formatNumber(responseMinutes)} min` : "skutečný dojezd není doložen"}.`,
      needsBaseReview
        ? `Dokud není doloženo vyhovující pokrytí do 20 minut, používá se scénářová kapacita 1 výjezdový vůz na ${formatNumber(populationPerVehicle)} obyvatel; nejméně ${formatNumber(vehiclesPerBase)} vozidla tvoří jednu základnu = ${responseVehicles} vozů.`
        : "Doložený dojezd do celého řešeného území je nejvýše 20 minut; nová základna se proto v bilanci nezapočítá, pokud kraj potvrdí pokrytí."
      ,
      coveredIsValid
        ? `Doplňková bilance obyvatel ve spádu: ${formatNumber(catchmentPopulation)} − ${formatNumber(covered)} = ${formatNumber(uncovered)} obyvatel bez doloženého pokrytí.`
        : `Bez doloženého počtu pokrytých obyvatel se za nepokrytou poptávku považuje celý spád ${formatNumber(uncovered)} obyvatel.`,
      "Plocha a typ území jsou uvedeny pro krajské/GIS posouzení; z plochy samotné nelze spolehlivě převést uliční síť na 20minutový dojezd."
    ], [
      "Zákon č. 374/2011 Sb., § 5, vyžaduje krajské plánování sítě ZZS a dostupnost do 20 minut od převzetí pokynu k výjezdu, s výjimkami dle zákona. Metodika MMR nestanoví populační kvótu základen. Poměr obyvatel na vůz a počet vozů na základnu jsou editovatelné provozní scénáře, nikoli norma; potvrzuje je krajský plán/poskytovatel.",
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], needsBaseReview ? responseVehicles : 0, [
      catchmentPopulation === null
        ? "Chybí počet obyvatel ve spádu; počet vozů nelze vypočítat."
        : `Modelový počet vozů = max(${formatNumber(vehiclesPerBase)} vozů/základnu, ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vůz)) = ${responseVehicles} vozů při nevyhovujícím nebo neověřeném dojezdu.`
    ]);
  }

  if (item.id === "VYB_16A") {
    const covered = getTargetedPlanInput(item, "HZS_COVERED_POPULATION");
    const responseMinutes = getTargetedPlanInput(item, "HZS_RESPONSE_MINUTES");
    const availableVehicles = getTargetedPlanInput(item, "HZS_AVAILABLE_VEHICLES");
    const populationPerVehicle = getTargetedPlanInput(item, "MODEL_POPULATION_PER_VEHICLE");
    const source = getTargetedPlanSource(item, "HZS_COVERAGE_SOURCE");
    const responseLimit = ["A", "B"].includes(territory) ? 7 : 15;
    const validResponse = Boolean(source) && responseMinutes !== null;
    const coveredIsValid = covered !== null && catchmentPopulation !== null && covered <= catchmentPopulation;
    const uncovered = coveredIsValid ? Math.max(0, catchmentPopulation - covered) : catchmentPopulation;
    const requiredVehicles = uncovered === null ? null : Math.ceil(uncovered / populationPerVehicle);
    const vehicleCapacityVerified = availableVehicles !== null && Number.isInteger(availableVehicles)
      && availableVehicles >= 0 && Boolean(source) && validResponse && responseMinutes <= responseLimit;
    const additionalVehicles = requiredVehicles === null ? null : vehicleCapacityVerified
      ? Math.max(0, requiredVehicles - availableVehicles) : requiredVehicles;
    const needsReview = !validResponse || responseMinutes > responseLimit
      || !vehicleCapacityVerified || additionalVehicles > 0;
    return estimate([
      { label: "Obyvatelé řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Obyvatelé včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Cílový interval dojezdu podle pracovního modelu", value: ["A", "B"].includes(territory) ? "5–7 min" : "10–15 min", unit: "" },
      { label: "Ověřená doba příjezdu", value: validResponse ? responseMinutes : null, unit: "min" },
      { label: "Obyvatelé pokrytí stávající jednotkou", value: coveredIsValid ? covered : null, unit: "obyvatel" },
      { label: "Obyvatelé bez doloženého pokrytí", value: uncovered, unit: "obyvatel" },
      { label: "Odhadovaný počet vozidel potřebný pro spád",
        value: requiredVehicles === null ? null : `${formatNumber(requiredVehicles)} ${czechCountUnit(requiredVehicles, "vozidlo", "vozidla", "vozidel")}`, unit: "" },
      { label: "Návrhový normativ vozidel (scénář)", value: populationPerVehicle, unit: "obyvatel/vozidlo" },
      { label: "Doplňkový počet vozidel k prověření",
        value: additionalVehicles === null ? null : `${formatNumber(additionalVehicles)} ${czechCountUnit(additionalVehicles, "vozidlo", "vozidla", "vozidel")}`, unit: "" },
      { label: "Výsledek posouzení", value: needsReview ? "Potřebu HZS započítat k ověření" : "Dojezd vyhovuje; potřeba se nezapočítá", unit: "" }
    ], [
      `Počet obyvatel ve spádu po záměru = ${catchmentPopulation === null ? "chybí vstup" : `${formatNumber(catchmentPopulation)} obyvatel`}; typ ${territory}; plocha ${context.areaHa === null ? "chybí vstup" : `${formatNumber(context.areaHa)} ha`}.`,
      validResponse
        ? `Ověřený dojezd ${formatNumber(responseMinutes)} min ${responseMinutes <= responseLimit ? "splňuje" : "nesplňuje"} horní mez ${responseLimit} min pro typ ${territory}.`
        : `Dojezd není doložen; pro typ ${territory} se používá horní mez ${responseLimit} min a potřeba zůstává započtena k ověření.`,
      vehicleCapacityVerified
        ? `Scénářová poptávka = ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${requiredVehicles}; po odečtu ${availableVehicles} ověřených vozidel zbývá ${additionalVehicles}.`
        : `Při absenci ověřeného volného vozidla se započítá celý scénář: ceil(${formatNumber(uncovered)} obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${requiredVehicles} vozidel k prověření.`,
      needsReview
        ? "Při chybějícím či nevyhovujícím dojezdu, nebo při doloženém nedostatku vozidel, zůstává potřeba započtena do posouzení."
        : "Doložený dojezd vyhovuje a zadaná kapacita vozidel nepřekračuje potřebu; nová stanice se v této bilanci nezapočítá."
    ], [
      "Metodika MMR/TB050MMR01, s. 12, odkazuje u hasičské stanice na dojezd podle přílohy k zákonu č. 133/1985 Sb.; pracovní list Cluster_Doprava_a_Bezpeci uvádí orientačně A/B 5–7 min a C/D 10–15 min. Poměr obyvatel na vozidlo, počet vozidel na stanici, plochy a cena jsou editovatelné scénáře, nikoli normy. Rozhoduje požární poplachový plán a HZS.",
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], additionalVehicles, [uncovered === null
      ? "Chybí počet obyvatel bez doloženého pokrytí; požadovaný počet vozidel nelze vypočítat."
      : `Modelová potřeba vozidel = ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${formatNumber(requiredVehicles)} vozidel.`]);
  }

  if (item.id === "VYB_17A") {
    const responseMinutes = getTargetedPlanInput(item, "POLICE_RESPONSE_MINUTES");
    const source = getTargetedPlanSource(item, "POLICE_COVERAGE_SOURCE");
    const localCoverage = ["A", "B"].includes(territory) ? getTargetedCoverageEstimate(600) : null;
    const responseValid = ["A", "B"].includes(territory)
      && Boolean(source) && responseMinutes !== null;
    const responseLimit = 10;
    const populationPerOfficer = getTargetedPlanInput(item, "MODEL_POPULATION_PER_OFFICER");
    const officersPerStation = getTargetedPlanInput(item, "MODEL_CAPACITY_PER_FACILITY");
    const modeledOfficers = catchmentPopulation === null || !populationPerOfficer ? null
      : Math.ceil(catchmentPopulation / populationPerOfficer);
    const modeledStations = modeledOfficers === null || !officersPerStation
      ? null : Math.ceil(modeledOfficers / officersPerStation);
    const proposedStations = responseValid && responseMinutes <= responseLimit
      ? 0 : modeledStations;
    const formatStations = (value) => value === null ? null
      : `${formatNumber(value)} ${czechCountUnit(value, "služebna", "služebny", "služeben")}`;
    return estimate([
      { label: "Počet obyvatel řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Počet obyvatel včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Scénářový počet policistů k zajištění", value: modeledOfficers, unit: "policistů" },
      { label: "Modelový počet služeben", value: formatStations(modeledStations), unit: "" },
      { label: "Modelová potřeba nových služeben po kontrole dojezdu", value: formatStations(proposedStations), unit: "" },
      { label: "Ověřená dojezdová doba hlídky (typ A/B)", value: responseValid ? responseMinutes : null, unit: "min" },
      { label: "Orientační body pěší sítě do 600 m", value: localCoverage?.count ?? null, unit: "bodů" }
    ], [
      `Počet policistů k prověření = ${catchmentPopulation === null ? "chybí počet obyvatel" : `${formatNumber(catchmentPopulation)} obyvatel ÷ ${formatNumber(populationPerOfficer)} obyvatel/policistu = ${formatNumber(modeledOfficers)} policistů`}; modelová služebna počítá nejvýše s ${formatNumber(officersPerStation)} policisty.`,
      `Typ ${territory}, plocha ${context.areaHa === null ? "chybí vstup" : `${formatNumber(context.areaHa)} ha`}, obyvatelé ve spádu ${catchmentPopulation === null ? "chybí vstup" : formatNumber(catchmentPopulation)}.`,
      responseValid
        ? `Ověřený dojezd hlídky ${formatNumber(responseMinutes)} min ${responseMinutes <= responseLimit ? "splňuje" : "nesplňuje"} orientační mez ${responseLimit} min.`
        : `Dojezd hlídky není doložen; modelové služebny se ponechávají k prověření. V typech C/D se dostupnost ověřuje v rámci ORP.`,
      proposedStations === 0
        ? "Vyhovující ověřený dojezd v typu A/B znamená nulovou modelovou potřebu nové služebny; personální potřeba se přesto uvádí pro kontrolu pokrytí."
        : `Pro návrhovou výměru se uvažuje ${formatNumber(proposedStations)} služeben, každá pro nejvýše ${formatNumber(officersPerStation)} policistů; nejde o závazný služební normativ.`
    ], [
      "Pracovní list Cluster_Doprava_a_Bezpeci uvádí pro A/B dojezd hlídky do 10 minut a pěší dostupnost 600 m; pro C/D obsluhu v rámci ORP. Zákon č. 273/2008 Sb. nestanoví kvótu policistů ani služeben na obyvatele. Poměr obyvatel na policistu a kapacita policistů na služebnu jsou editovatelné scénářové předpoklady, nikoli oficiální český personální normativ."
    ], proposedStations === null || modeledOfficers === null ? null : proposedStations === 0 ? 0 : modeledOfficers, [
      catchmentPopulation === null || !populationPerOfficer
        ? "Chybí počet obyvatel nebo poměr obyvatel na policistu; personální potřebu nelze vypočítat."
        : `Modelová personální potřeba = ceil(${formatNumber(catchmentPopulation)} obyvatel ÷ ${formatNumber(populationPerOfficer)} obyvatel/policistu) = ${formatNumber(modeledOfficers)} policistů; orientační počet služeben = ceil(${formatNumber(modeledOfficers)} ÷ ${formatNumber(officersPerStation)} policistů/služebnu) = ${formatNumber(modeledStations)}.`
    ]);
  }

  if (item.id === "VYB_17B") {
    const enteredDemand = getTargetedPlanInput(item, "SHELTER_PLANNED_DEMAND");
    const populationShare = getTargetedPlanInput(item, "MODEL_POPULATION_SHARE");
    const populationBasis = catchmentPopulation ?? projectPopulation;
    const demand = enteredDemand ?? (populationBasis === null || populationShare === null
      ? null : Math.ceil(populationBasis * populationShare));
    const capacity = getTargetedPlanInput(item, "SHELTER_EXISTING_CAPACITY");
    const distance = getTargetedPlanInput(item, "SHELTER_ACCESS_DISTANCE");
    const source = getTargetedPlanSource(item, "SHELTER_VERIFICATION_SOURCE");
    const valid = Boolean(source) && demand !== null
      && capacity !== null && Number.isInteger(capacity) && distance !== null;
    const usableCapacity = valid && distance <= 500 ? capacity : valid ? 0 : null;
    const deficit = demand === null ? null : valid ? Math.max(0, demand - usableCapacity) : demand;
    const coverage = getTargetedCoverageEstimate(500);
    return estimate([
      { label: enteredDemand === null
        ? `Poptávka po úkrytu (${populationShare === null ? "nevyplněno" : `${formatNumber(populationShare * 100)} %`} obyvatel ve spádu)`
        : "Potřeba osob podle místního plánu", value: demand, unit: "osob" },
      { label: "Počet obyvatel včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Ověřená kapacita stávajících úkrytů", value: valid ? capacity : null, unit: "osob" },
      { label: "Vzdálenost k započtenému úkrytu", value: distance, unit: "m" },
      { label: "Kapacita započtená do 500 m", value: usableCapacity, unit: "osob" },
      { label: "Poptávka nepokrytá ověřenou kapacitou úkrytů", value: deficit, unit: "osob" },
      { label: "Orientační body pěší sítě do 500 m", value: coverage?.count ?? null, unit: "bodů" }
    ], valid ? [
      `Započtená kapacita = ${formatNumber(distance)} m ${distance <= 500 ? "≤" : ">"} 500 m, tedy ${formatNumber(usableCapacity)} osob.`,
      `Poptávka nepokrytá ověřenou kapacitou = max(0, ${formatNumber(demand)} osob − ${formatNumber(usableCapacity)} osob) = ${formatNumber(deficit)} osob.`,
      coverage
        ? `Plošná kontrola dosahu: ${formatNumber(context.areaHa)} ha, typ ${territory}, 500 m pěšky; mřížka ${coverage.columns} × ${coverage.rows} = ${coverage.count} bodů.`
        : "Plošnou kontrolu 500 m nelze odhadnout bez plochy, tvaru a koeficientu pěší sítě.",
      `Návrhová kapacita se převádí na podzemní stavbu podle upravitelných hodnot ${formatNumber(getTargetedPlanInput(item, "MODEL_AREA_PER_UNIT"))} m²/osobu a ${formatNumber(getTargetedPlanInput(item, "MODEL_CONSTRUCTION_HEIGHT"))} m konstrukční výšky.`
    ] : [
      `Projekční poptávka = ${demand === null ? "chybí počet obyvatel" : formatNumber(demand)} osob; bez ověření kapacity se pro návrh použije celý tento scénář, nikoli potvrzený deficit.`
    ], [
      "Vyhláška č. 380/2002 Sb., § 22 odst. 2, stanoví dostupnost do 500 m od míst velké koncentrace osob; neurčuje univerzální kapacitu pro každého obyvatele. Výchozí horní scénář počítá s obyvateli ve spádu, podíl lze upravit podle místního plánu. RTS 2026 neuvádí cenu stálého úkrytu: JKSO 814.2 se týká pozemních nádrží a 815.9 má jednotku délky, proto sazba 30 000 Kč/m³ zůstává transparentním editovatelným odhadem mimo RTS. Nezahrnuje technologie ochranného provozu.",
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], deficit, [demand === null
      ? "Chybí počet obyvatel nebo podíl populace určený k ukrytí."
      : enteredDemand === null
        ? `Scénářová potřeba = ceil(${formatNumber(populationBasis)} obyvatel ve spádu × ${formatNumber(populationShare * 100)} %) = ${formatNumber(demand)} osob.`
        : `Potřeba podle místního plánu = ${formatNumber(demand)} osob.`]);
  }

  return null;
}

function getTargetedPlanSummaryText(item, estimate) {
  const metrics = estimate?.metrics ?? [];
  const metricValue = (label) => metrics.find((metric) => metric.label === label)?.value;
  if (["VYB_02C_1", "VYB_02C_2"].includes(item.id)) {
    const unmetSeats = metricValue("Poptávka po odečtu ověřené volné kapacity");
    const verifiedFreeCapacity = metricValue("Ověřená volná kapacita") !== null
      && metricValue("Ověřená volná kapacita") !== undefined;
    const demand = metricValue("Odhadovaná poptávka po místech");
    if (unmetSeats !== null && unmetSeats !== undefined)
      return verifiedFreeCapacity
        ? `${formatNumber(unmetSeats)} míst k doplnění`
        : `${formatNumber(unmetSeats)} míst poptávky (ověřit kapacitu)`;
    if (demand !== null && demand !== undefined)
      return `${formatNumber(demand)} žáků (scénář)`;
  }
  const status = metricValue("Výsledek podmíněného posouzení")
    ?? metricValue("Výsledek posouzení")
    ?? metricValue("Přítomnost v obci podle typu a populace");
  if (status !== null && status !== undefined) return String(status);
  if (item.id === "VYB_01_J") {
    const children = metrics.find((metric) => metric.label.startsWith("Odhad poptávky po péči"))?.value;
    return children === null || children === undefined
      ? "ověřit demografii"
      : `${formatNumber(children, 1)} dětí (indikativně)`;
  }
  if (item.id === "VYB_09A_4") {
    const fte = metricValue("Poptávka po gynekologii k ověření / po odečtu kapacity")
      ?? metricValue("Indikativní personální ekvivalent");
    return fte === null || fte === undefined
      ? "ověřit kapacitu"
      : `${formatNumber(fte, 2)} souč. zátěže`;
  }
  if (item.id === "VYB_17A") {
    const stations = metricValue("Modelová potřeba nových služeben po kontrole dojezdu");
    return stations === null || stations === undefined
      ? "ověřit pokrytí"
      : `${formatNumber(stations)} modelových služeben`;
  }
  if (item.id === "VYB_17B") {
    const deficit = metricValue("Poptávka nepokrytá ověřenou kapacitou úkrytů");
    return deficit === null || deficit === undefined
      ? "ověřit místní plán"
      : `${formatNumber(deficit)} osob deficitu`;
  }
  return "metodické ověření";
}

function getProjectPlanningEstimate(item, kind) {
  if (isResidentParkingItem(item)) return null;
  if (Object.hasOwn(TARGETED_PLAN_INPUTS, item.id)
    || ["VYB_02C_1", "VYB_02C_2", "VYB_09A_4"].includes(item.id)) {
    return getTargetedPlanEstimate(item);
  }
  if (item.id === "VYB_04A" && kind === "legal") {
    const qualifyingAreaHa = Math.max(0, numericOrNull(getPrimary("S_zastav")) ?? 0);
    return {
      type: "public-space",
      qualifyingAreaHa,
      minimumArea: Math.ceil(qualifyingAreaHa * 500),
      unit: "m²"
    };
  }
  if (item.id === HEATING_BOILER_ID && kind === "infrastructure") {
    const energy = getAggregateProjectEnergyMetrics(item);
    const heatingShare = numericOrNull(getItemEstimateValue(item, "BOILER_HEATING_SHARE"));
    const fullLoadHours = numericOrNull(getItemEstimateValue(item, "BOILER_FULL_LOAD_HOURS"));
    const reserveFactor = numericOrNull(getItemEstimateValue(item, "BOILER_DESIGN_RESERVE"));
    const unitCapacity = numericOrNull(getItemEstimateValue(item, "BOILER_UNIT_CAPACITY_KW"));
    const heatDensity = numericOrNull(getItemEstimateValue(item, "BOILER_HEAT_DENSITY_KWH_M2_Y"));
    const unitCost = numericOrNull(getItemEstimateValue(item, "BOILER_UNIT_COST"));
    const surcharge = numericOrNull(getItemEstimateValue(item, "ALT_COST_SURCHARGE"));
    if ([heatingShare, fullLoadHours, reserveFactor, unitCapacity, heatDensity, surcharge]
      .some((value) => value === null) || fullLoadHours <= 0 || unitCapacity <= 0) return null;
    const annualHeatDemand = (energy.modeledNonElectricEnergy + energy.fallbackHeatDemand) * heatingShare;
    const designCapacity = annualHeatDemand / fullLoadHours * reserveFactor;
    const pointCount = designCapacity > 0 ? Math.ceil(designCapacity / unitCapacity) : 0;
    const pointCost = unitCost === null ? null : Math.round(pointCount * unitCost);
    const zrn = pointCost;
    return {
      type: "infrastructure",
      isHeatingBoiler: true,
      annualHeatDemand,
      designCapacity,
      quantity: designCapacity,
      unit: "kW tepelného výkonu",
      pointCount,
      pointRate: unitCost,
      pointCost,
      pointLabel: "Kotlový modul",
      annualElectricity: energy.annualElectricity,
      peakElectricalLoad: energy.peakElectricalLoad,
      modeledAnnualElectricity: energy.modeledAnnualElectricity,
      modeledNonElectricEnergy: energy.modeledNonElectricEnergy,
      modeledEquipmentCount: energy.modeledEquipmentCount,
      fallbackAnnualElectricity: energy.fallbackAnnualElectricity,
      energyDensity: numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_DENSITY_W_M2")),
      peakToAverageRatio: numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO")),
      fallbackFloorArea: energy.fallbackFloorArea,
      fallbackHeatingArea: energy.fallbackHeatingArea,
      fallbackHeatDemand: energy.fallbackHeatDemand,
      heatDensity,
      projectFloorArea: energy.projectFloorArea,
      fullLoadHours,
      reserveFactor,
      heatingShare,
      unitCapacity,
      costSurcharge: surcharge,
      zrn,
      cost: zrn === null ? null : Math.round(zrn * (1 + surcharge))
    };
  }
  if (TECHNICAL_ALTERNATIVE_IDS.has(item.id) && kind === "infrastructure") {
    const footprint = getAggregateProjectSurfaceMetrics();
    const realizationPlan = getRealizationPlan();
    const population = Math.max(0, numericOrNull(getPrimary("O_celk")) ?? 0);
    const rainfallDepth = numericOrNull(getItemEstimateValue(item, "ALT_RAINFALL_DEPTH_MM"));
    const runoffCoefficient = numericOrNull(getItemEstimateValue(item, "ALT_RUNOFF_COEFFICIENT"));
    const waterUse = numericOrNull(getItemEstimateValue(item, "ALT_WATER_LITERS_PER_PERSON_DAY"));
    const peakFactor = numericOrNull(getItemEstimateValue(item, "ALT_PEAK_DAY_FACTOR"));
    const wastewaterReturn = numericOrNull(getItemEstimateValue(item, "ALT_WASTEWATER_RETURN"));
    const networkLength = numericOrNull(getItemEstimateValue(item, "ALT_NETWORK_LENGTH_M"));
    const waterPipeRate = numericOrNull(getItemEstimateValue(item, "ALT_WATER_PIPE_RATE"));
    const waterPointCount = numericOrNull(getItemEstimateValue(item, "ALT_WATER_METER_CHAMBER_COUNT"));
    const waterPointRate = numericOrNull(getItemEstimateValue(item, "ALT_WATER_METER_CHAMBER_RATE"));
    const sewerPipeRate = numericOrNull(getItemEstimateValue(item, "ALT_SEWER_PIPE_RATE"));
    const sewerManholeSpacing = numericOrNull(getItemEstimateValue(item, "ALT_SEWER_MANHOLE_SPACING_M"));
    const sewerManholeRate = numericOrNull(getItemEstimateValue(item, "ALT_SEWER_MANHOLE_RATE"));
    const rainwaterRate = numericOrNull(getItemEstimateValue(item, "ALT_RAINWATER_RATE"));
    const energyDensity = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_DENSITY_W_M2"));
    const peakToAverageRatio = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO"));
    const powerCableRate = numericOrNull(getItemEstimateValue(item, "ALT_POWER_CABLE_RATE"));
    const transformerCapacity = numericOrNull(getItemEstimateValue(item, "ALT_TRANSFORMER_CAPACITY_KW"));
    const transformerRate = numericOrNull(getItemEstimateValue(item, "ALT_TRANSFORMER_RATE"));
    const surcharge = numericOrNull(getItemEstimateValue(item, "ALT_COST_SURCHARGE"));
    if ([rainfallDepth, runoffCoefficient, waterUse, peakFactor, wastewaterReturn,
      networkLength, waterPipeRate, waterPointCount, waterPointRate, sewerPipeRate,
      sewerManholeSpacing, sewerManholeRate, rainwaterRate, energyDensity, powerCableRate,
      transformerCapacity, transformerRate, surcharge, peakToAverageRatio]
      .some((value) => value === null)) {
      return null;
    }

    let metric;
    let quantity;
    let unit;
    let lineRate;
    let lineCost;
    let pointCount = 0;
    let pointRate = 0;
    let pointLabel = "";
    let energyMetrics = null;
    if (item.id === "VYB_20A") {
      metric = footprint.runoffArea * rainfallDepth / 1000 * runoffCoefficient;
      quantity = metric;
      unit = "m³ retenčního objemu";
      lineRate = rainwaterRate;
      lineCost = quantity * lineRate;
      pointLabel = "Nádrž a stavební práce";
    } else if (item.id === "VYB_20B_1") {
      metric = population * waterUse / 1000;
      quantity = metric * peakFactor;
      unit = "m³/den návrhové kapacity";
      lineRate = waterPipeRate;
      lineCost = networkLength * lineRate;
      pointCount = waterPointCount;
      pointRate = waterPointRate;
      pointLabel = "Vodoměrná a armaturní šachta";
    } else if (item.id === "VYB_20B_2") {
      metric = population * waterUse / 1000 * wastewaterReturn;
      quantity = metric * peakFactor;
      unit = "m³/den návrhové kapacity";
      lineRate = sewerPipeRate;
      lineCost = networkLength * lineRate;
      pointCount = networkLength > 0 ? Math.ceil(networkLength / sewerManholeSpacing) + 1 : 0;
      pointRate = sewerManholeRate;
      pointLabel = "Revizní šachty";
    } else {
      energyMetrics = getAggregateProjectEnergyMetrics(item);
      metric = energyMetrics.peakElectricalLoad;
      quantity = metric;
      unit = "kW návrhového příkonu";
      lineRate = powerCableRate;
      lineCost = networkLength * lineRate;
      pointCount = quantity > 0 ? Math.ceil(quantity / transformerCapacity) : 0;
      pointRate = transformerRate;
      pointLabel = "Trafostanice";
    }
    const pointCost = pointCount * pointRate;
    const zrn = Math.round(lineCost + pointCost);
    return {
      type: "infrastructure",
      population,
      runoffArea: footprint.runoffArea,
      totalFloorArea: footprint.totalFloorArea,
      averageDailyWater: population * waterUse / 1000,
      averageDailyWastewater: population * waterUse / 1000 * wastewaterReturn,
      annualElectricity: energyMetrics?.annualElectricity ?? null,
      modeledAnnualElectricity: energyMetrics?.modeledAnnualElectricity ?? null,
      modeledPeakElectricalLoad: energyMetrics?.modeledPeakElectricalLoad ?? null,
      modeledEquipmentCount: energyMetrics?.modeledEquipmentCount ?? 0,
      fallbackFloorArea: energyMetrics?.fallbackFloorArea ?? 0,
      fallbackPeakElectricalLoad: energyMetrics?.fallbackPeakElectricalLoad ?? 0,
      peakToAverageRatio,
      sourceMetric: metric,
      quantity,
      unit,
      networkLength,
      lineRate,
      lineCost: Math.round(lineCost),
      pointCount,
      pointRate,
      pointLabel,
      pointCost: Math.round(pointCost),
      costSurcharge: surcharge,
      zrn,
      cost: Math.round(zrn * (1 + surcharge)),
      rainfallDepth,
      runoffCoefficient,
      waterUse,
      peakFactor,
      wastewaterReturn,
      energyDensity,
      transformerCapacity
    };
  }

  if (item.id === "VYB_20"
    && String(item.category ?? "").toLocaleUpperCase("cs").includes("ZDRAVOTNICTVÍ")
    && kind === "recommended") {
    const population = numericOrNull(derivedValues().O_prah);
    const peoplePerCentre = getProjectEstimateValue(item, "POLY_OBYV_NA_CENTRUM");
    const usableAreaPerCentre = getProjectEstimateValue(item, "POLY_UZITNA_PLOCHA");
    const costPerCubicMeter = getProjectEstimateValue(item, "POLY_CENA_M3");
    const costSurcharge = getProjectEstimateValue(item, "POLY_PRIRAZKA");
    const fixedCost = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_FIXED_COST"));
    const totalFloorFactor = getProjectEstimateValue(item, "K_CELKOVA_PLOCHA");
    const volumeFactor = getProjectEstimateValue(item, "K_OBESTAVENY");
    const siteAreaFactor = getProjectEstimateValue(item, "K_PLOCHA_AREALU");
    const footprintFactor = getProjectEstimateValue(item, "K_ZASTAVENI");
    const pavedShare = getProjectEstimateValue(item, "K_ZPEVNENI");
    if ([population, peoplePerCentre, usableAreaPerCentre, costSurcharge, fixedCost]
      .some((value) => value === null)
      || [totalFloorFactor, volumeFactor, siteAreaFactor, footprintFactor, pavedShare]
        .some((value) => value === null)
      || peoplePerCentre <= 0) return null;

    const centerEquivalent = population / peoplePerCentre;
    const usableArea = centerEquivalent * usableAreaPerCentre;
    const totalFloorArea = usableArea * totalFloorFactor;
    const volume = Math.round(totalFloorArea * volumeFactor);
    const siteArea = Math.round(totalFloorArea * siteAreaFactor);
    const footprintArea = Math.round(totalFloorArea * footprintFactor);
    const pavedArea = Math.round(Math.max(0, siteArea - footprintArea) * pavedShare);
    const zrn = costPerCubicMeter === null
      ? null : Math.round(volume * costPerCubicMeter + fixedCost);
    return {
      type: "polyclinic",
      population,
      peoplePerCentre,
      centerEquivalent,
      centreTriggered: population >= peoplePerCentre,
      usableArea,
      totalFloorArea,
      volume,
      siteArea,
      footprintArea,
      pavedArea,
      zrn,
      cost: zrn === null ? null : Math.round(zrn * (1 + costSurcharge)),
      costPerCubicMeter,
      fixedCost,
      costSurcharge,
      totalFloorFactor,
      volumeFactor,
      siteAreaFactor,
      footprintFactor,
      pavedShare,
      usableAreaPerCentre
    };
  }

  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
  const baseUsableArea = numericOrNull(item.usableArea);
  if (baseUsableArea === null) return null;
  const totalFloorFactor = getProjectEstimateValue(item, "K_CELKOVA_PLOCHA");
  const volumeFactor = numericOrNull(getItemEstimateValue(item, "ITEM_VOLUME_PER_AREA"))
    ?? getProjectEstimateValue(item, "K_OBESTAVENY");
  const siteAreaFactor = getProjectEstimateValue(item, "K_PLOCHA_AREALU");
  const footprintFactor = getProjectEstimateValue(item, "K_ZASTAVENI");
  const pavedShare = getProjectEstimateValue(item, "K_ZPEVNENI");
  const costRate = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_RATE"));
  const costRateBasis = getItemEstimateValue(item, "ITEM_RTS_RATE_BASIS");
  const fixedCost = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_FIXED_COST"));
  const costSurcharge = numericOrNull(getItemEstimateValue(item, "ITEM_COST_SURCHARGE",
    item.rtsPricing?.surcharge ?? 0));
  if ([totalFloorFactor, volumeFactor, siteAreaFactor, footprintFactor, pavedShare, fixedCost, costSurcharge]
    .some((value) => value === null)) return null;
  const scale = getStandardMetricScale(item, kind);
  const standardUsableArea = getStandardMetricValue(baseUsableArea, scale);
  const upliftPercent = kind === "recommended" ? getSchoolRecommendationPercent(item) : 0;
  const reserveUsableArea = standardUsableArea * upliftPercent / 100;
  const usableArea = standardUsableArea + reserveUsableArea;
  const totalFloorArea = usableArea * totalFloorFactor;
  const volume = Math.round(totalFloorArea * volumeFactor);
  const siteArea = Math.round(totalFloorArea * siteAreaFactor);
  const footprintArea = Math.round(totalFloorArea * footprintFactor);
  const pavedArea = Math.round(Math.max(0, siteArea - footprintArea) * pavedShare);
  const costBasis = costRateBasis === "volume" ? volume : getItemEstimateQuantity(item) * scale;
  const zrn = costRate === null || costBasis === null ? null
    : Math.round(costBasis * costRate + fixedCost);
  return {
    type: "school",
    baseUsableArea,
    standardUsableArea,
    upliftPercent,
    reserveUsableArea,
    usableArea,
    totalFloorArea,
    volume,
    siteArea,
    footprintArea,
    pavedArea,
    costPerCubicMeter: costRateBasis === "volume" ? costRate : null,
    costRate,
    costRateBasis,
    costBasis,
    fixedCost,
    zrn,
    costSurcharge,
    cost: zrn === null ? null : Math.round(zrn * (1 + costSurcharge)),
    totalFloorFactor,
    volumeFactor,
    siteAreaFactor,
    footprintFactor,
    pavedShare
  };
  }

  if (item.id === "VYB_07B" || item.manualCoverageCheck) return null;
  const quantity = getItemEstimateQuantity(item);
  const usableAreaPerUnit = numericOrNull(getItemEstimateValue(item, "ITEM_USABLE_AREA_PER_UNIT"));
  const totalFloorFactor = numericOrNull(getItemEstimateValue(item, "ITEM_TOTAL_FLOOR_FACTOR"));
  const volumePerArea = numericOrNull(getItemEstimateValue(item, "ITEM_VOLUME_PER_AREA"));
  const siteAreaPerUnit = numericOrNull(getItemEstimateValue(item, "ITEM_SITE_AREA_PER_UNIT"));
  const pavedAreaPerUnit = numericOrNull(getItemEstimateValue(item, "ITEM_PAVED_AREA_PER_UNIT"));
  const rate = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_RATE"));
  const rateBasis = getItemEstimateValue(item, "ITEM_RTS_RATE_BASIS");
  const fixedCost = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_FIXED_COST"));
  const surcharge = numericOrNull(getItemEstimateValue(item, "ITEM_COST_SURCHARGE"));
  const usableArea = usableAreaPerUnit === null ? null : quantity * usableAreaPerUnit;
  const totalFloorArea = usableArea === null || totalFloorFactor === null
    ? null : usableArea * totalFloorFactor;
  const volume = totalFloorArea === null || volumePerArea === null
    ? null : Math.round(totalFloorArea * volumePerArea);
  const siteArea = siteAreaPerUnit === null ? null : Math.round(quantity * siteAreaPerUnit);
  const pavedArea = pavedAreaPerUnit === null ? null : Math.round(quantity * pavedAreaPerUnit);
  const costBasis = rateBasis === "volume" ? volume : quantity;
  const zrn = rate === null || costBasis === null || fixedCost === null
    ? null : Math.round(costBasis * rate + fixedCost);
  const savedRate = state.itemEstimateValues[item.key]?.ITEM_RTS_RATE;
  const rtsConstructionRate = getRtsConstructionRate(item);
  let rateSource;
  if (item.id === "VYB_15B_2"
    || (!getRtsBuildingGroup(item) && savedRate !== undefined && savedRate !== null)) {
    rateSource = "user-estimate";
  } else if (rtsConstructionRate !== undefined) {
    rateSource = rtsConstructionRate === null ? "missing" : "RTS";
  } else if (item.id === "VYB_04B") {
    rateSource = "scenario-estimate";
  } else if (item.id === "VYB_10B_1") {
    rateSource = "scenario-estimate";
  } else if (item.rtsPricing?.rate) {
    rateSource = "RTS";
  } else if (numericOrNull(item.unitPrice) > 0) {
    rateSource = "workbook-unit-cost";
  } else if (ITEM_ESTIMATE_SCENARIOS[item.id]) {
    rateSource = "scenario-estimate";
  } else if (numericOrNull(item.cost) > 0 && getItemEstimateDefaultQuantity(item) > 0) {
    rateSource = "workbook-total-cost-derived";
  } else {
    rateSource = estimateComparableUnitRate(item) !== null ? "comparable-estimate" : "missing";
  }
  return {
    type: "item",
    quantity,
    quantityUnit: item.unit,
    usableAreaPerUnit,
    usableArea,
    totalFloorFactor,
    totalFloorArea,
    volumePerArea,
    volume,
    siteAreaPerUnit,
    siteArea,
    pavedAreaPerUnit,
    pavedArea,
    rate,
    rateSource,
    rateBasis,
    fixedCost,
    surcharge,
    zrn,
    cost: zrn === null || surcharge === null ? null : Math.round(zrn * (1 + surcharge))
  };
}

function getHousingContributionEstimate(item) {
  if (item.id !== "VYB_07B") return null;
  const totalUnits = numericOrNull(derivedValues().BJ_pocet);
  const apartmentShare = getResidentialMixAssumptions().apartmentShare;
  const totalApartments = totalUnits === null ? null : Math.round(totalUnits * apartmentShare);
  const buyoutShare = numericOrNull(getPrimary("k_odkup_obec"));
  const lostMarginShare = numericOrNull(getPrimary("m_marze_soc"));
  const areaPerHome = numericOrNull(getSecondary("BYTY_PLOCHA_NA_BYT"));
  const volumePerArea = numericOrNull(getSecondary("BYTY_OBESTAVENI_NA_M2"));
  const valuePerCubicMeter = numericOrNull(getSecondary("BYTY_CENA_M3"));
  if ([totalUnits, totalApartments, buyoutShare, lostMarginShare, areaPerHome, volumePerArea, valuePerCubicMeter]
    .some((value) => value === null)) return null;

  const municipalApartments = Math.round(totalApartments * buyoutShare);
  const portfolioArea = totalApartments * areaPerHome;
  const municipalArea = municipalApartments * areaPerHome;
  const portfolioVolume = Math.round(portfolioArea * volumePerArea);
  const municipalVolume = Math.round(municipalArea * volumePerArea);
  const portfolioValue = portfolioVolume * valuePerCubicMeter;
  const municipalValue = municipalVolume * valuePerCubicMeter;
  return {
    totalUnits,
    totalApartments,
    apartmentShare,
    buyoutShare,
    municipalApartments,
    areaPerHome,
    portfolioArea,
    municipalArea,
    volumePerArea,
    portfolioVolume,
    municipalVolume,
    valuePerCubicMeter,
    portfolioValue,
    municipalValue,
    lostMarginShare,
    lostMargin: Math.round(municipalValue * lostMarginShare)
  };
}

function renderProjectEstimateControls(item, kind, estimate, targetedInputScope = "model") {
  let fields = [];
  let note = "";
  if (estimate?.type === "targeted") {
    fields = (TARGETED_PLAN_INPUTS[item.id] || []).filter((field) =>
      isTargetedModelInput(field) === (targetedInputScope === "model")
      && !(targetedInputScope === "model" && field.code === "MODEL_COST_PER_M3"
        && getRtsBuildingGroup(item)));
    note = targetedInputScope === "model"
      ? getRtsBuildingGroup(item)
        ? `Zvolená konstrukční charakteristika určuje sazbu RTS ${getRtsBuildingGroup(item)} v Kč/m³ OP; změna materiálu okamžitě přepočítá investici. ${item.id === "VYB_02C_2" ? "JKSO 801.4 je pro ZUŠ pouze cenový proxy." : ""} Ostatní plošné a kapacitní hodnoty jsou editovatelné scénáře.`
        : "Upravitelné vstupy návrhové kapacity, výměr a ceny. Výchozí hodnoty jsou scénářové, pokud není výslovně uveden zdroj RTS; nejde o projektový rozpočet."
      : "Zde se ověřuje skutečná okolní kapacita, dostupnost a zdroj. Ověření je oddělené od návrhové poptávky.";
  } else if (estimate?.type === "school") {
    fields = [
      ...PROJECT_ESTIMATE_INPUTS.filter((field) => field.code.startsWith("K_")
        && field.code !== "K_CENA_SKOLY" && field.code !== "K_OBESTAVENY"),
      ...ITEM_ESTIMATE_INPUTS.filter((field) => [
        "ITEM_VOLUME_PER_AREA", "ITEM_RTS_RATE", "ITEM_RTS_RATE_BASIS",
        "ITEM_RTS_FIXED_COST", "ITEM_COST_SURCHARGE"
      ].includes(field.code)).filter((field) =>
        !getRtsBuildingGroup(item) || !["ITEM_RTS_RATE", "ITEM_RTS_RATE_BASIS"].includes(field.code))
    ];
    note = "Předpoklady se ukládají samostatně pro tento školní objekt. Sazba a fixní náklady se přebírají z RTS modelu, pokud jsou v sešitu uvedeny.";
  } else if (estimate?.type === "polyclinic") {
    fields = [
      ...PROJECT_ESTIMATE_INPUTS.filter((field) =>
        !getRtsBuildingGroup(item) || field.code !== "POLY_CENA_M3"),
      ITEM_ESTIMATE_INPUTS.find((field) => field.code === "ITEM_RTS_FIXED_COST")
    ];
    note = "Předpoklady se ukládají samostatně pro tento odhad polikliniky. Nejde o hlavní demografické vstupy.";
  } else if (estimate?.type === "infrastructure") {
    fields = ALTERNATIVE_ESTIMATE_INPUTS.filter((field) => {
      if (item.id === HEATING_BOILER_ID) return [
        "ALT_ENERGY_DENSITY_W_M2", "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO",
        "BOILER_HEAT_DENSITY_KWH_M2_Y", "BOILER_HEATING_SHARE",
        "BOILER_FULL_LOAD_HOURS", "BOILER_DESIGN_RESERVE",
        "BOILER_UNIT_CAPACITY_KW", "BOILER_UNIT_COST", "ALT_COST_SURCHARGE"
      ].includes(field.code);
      if (item.id === "VYB_20A") return [
        "ALT_RAINFALL_DEPTH_MM", "ALT_RUNOFF_COEFFICIENT", "ALT_RAINWATER_RATE", "ALT_COST_SURCHARGE"
      ].includes(field.code);
      if (item.id === "VYB_20B_1") return [
        "ALT_WATER_LITERS_PER_PERSON_DAY", "ALT_PEAK_DAY_FACTOR", "ALT_NETWORK_LENGTH_M",
        "ALT_WATER_PIPE_RATE", "ALT_WATER_METER_CHAMBER_COUNT", "ALT_WATER_METER_CHAMBER_RATE",
        "ALT_COST_SURCHARGE"
      ].includes(field.code);
      if (item.id === "VYB_20B_2") return [
        "ALT_WATER_LITERS_PER_PERSON_DAY", "ALT_PEAK_DAY_FACTOR", "ALT_WASTEWATER_RETURN",
        "ALT_NETWORK_LENGTH_M", "ALT_SEWER_PIPE_RATE", "ALT_SEWER_MANHOLE_SPACING_M",
        "ALT_SEWER_MANHOLE_RATE", "ALT_COST_SURCHARGE"
      ].includes(field.code);
      return [
        "ALT_ENERGY_DENSITY_W_M2", "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO",
        "ALT_NETWORK_LENGTH_M", "ALT_POWER_CABLE_RATE",
        "ALT_TRANSFORMER_CAPACITY_KW", "ALT_TRANSFORMER_RATE", "ALT_COST_SURCHARGE"
      ].includes(field.code);
    });
    note = item.id === HEATING_BOILER_ID
      ? "Jde o scénářový odhad, nikoli projektový výpočet. Využívá roční spotřebu elektřiny a příkon modelovaných položek ze sešitu; chybějící údaje mají oddělené, editovatelné odhady. Cena modulu je orientačně odvozena sazbou 10 000 Kč/kW z jeho zadaného výkonu (výchozích 500 kW = 5 mil. Kč/modul); nejde o sazbu RTS a nezahrnuje projekt, kotelnu ani související technologie. Neelektrická energie není automaticky totožná s potřebou vytápění. Upravte předpoklady podle energetického posudku."
      : "Odhad mimo RTS a projektový rozpočet. Trasa je výchozím předpokladem 1,5 km. Sazby za metr jsou odvozeny z rozdělení nákladového scénáře TECH-COST-001 na tuto délku; šachty a trafostanice jsou samostatné pracovní odhady. Spotřeba domácností vychází z ČSÚ za rok 2023 (86,7, zaokrouhleno na 87 l/os./den). Všechny sazby a výměry upravte podle podmínek správců sítí a skutečné trasy.";
  } else if (item.id === "VYB_07B") {
    if (kind !== "recommended") return null;
    fields = HOUSING_CONTRIBUTION_INPUTS;
    note = "Editovatelné scénářové koeficienty; podíl městského odkupu a ušlá marže se nastavují v hlavních vstupech.";
  } else if (estimate?.type === "item") {
    fields = ITEM_ESTIMATE_INPUTS.filter((field) =>
      !getRtsBuildingGroup(item) || !["ITEM_RTS_RATE", "ITEM_RTS_RATE_BASIS"].includes(field.code));
    note = estimate.rateSource === "RTS"
      ? "Předpoklady se ukládají pro tuto položku. Cena vychází z RTS; chybějící výměrové koeficienty jsou pracovní odhady."
      : estimate.rateSource === "missing"
        ? "Předpoklady se ukládají pro tuto položku. Cena není v RTS ani v porovnatelných datech; sazbu lze zadat jako vlastní odhad mimo RTS."
        : estimate.rateSource === "scenario-estimate"
            ? item.id === "VYB_04B"
              ? "RTS v dostupném podkladu cenu veřejné parkové zeleně neuvádí. Výchozí sazba 2 250 000 Kč/park (4 500 m²) je převzata z pracovního scénáře zeleně, tedy 500 Kč/m² před přirážkou 23 %; jde o orientační odhad mimo RTS, upravte jej podle rozsahu a místních cen."
              : "Předpoklady se ukládají pro tuto položku. Výchozí cena i výměrové koeficienty jsou orientační odhady mimo RTS; upravte je podle záměru a místních cen."
        : "Předpoklady se ukládají pro tuto položku. Cena je orientační cenový odhad mimo RTS; ověřte ji podle nabídky nebo projektu.";
  }
  if (!fields.length) return null;

  const disclosure = createElement("details", "project-estimate-controls");
  addDisclosureSummary(disclosure, estimate?.type === "targeted"
    ? targetedInputScope === "model" ? "Upravit předpoklady návrhového výpočtu"
      : "Upravit údaje pro ověření okolní vybavenosti"
    : "Upravit předpoklady tohoto propočtu");
  disclosure.append(createElement("p", "project-estimate-input-note", note));
  const grid = createElement("div", "secondary-fields project-estimate-inputs");
  const suffix = `${item.id}-${kind}`;
  for (const field of fields) {
    if (item.id === "VYB_07B" && estimate?.type !== "school" && estimate?.type !== "polyclinic") {
      grid.append(createInputControl(field, getSecondary(field.code), true, suffix));
      continue;
    }
    const value = estimate?.type === "item" || estimate?.type === "infrastructure"
      ? getItemEstimateValue(item, field.code)
      : field.code.startsWith("K_") || field.code.startsWith("POLY_")
        ? getProjectEstimateValue(item, field.code)
        : getItemEstimateValue(item, field.code);
    const itemField = estimate?.type === "targeted"
      ? { ...field, color: TARGETED_FIELD_COLORS[field.code] || field.color }
      : ["item", "school"].includes(estimate?.type) && field.code === "ITEM_RTS_RATE"
      ? { ...field, label: estimate.type === "school" || estimate.rateSource === "RTS"
        ? "Cena RTS dle měrné jednotky" : "Odhad ceny mimo RTS" }
      : field;
    grid.append(createInputControl(itemField, value, true, suffix, "itemEstimate", item.key));
  }
  disclosure.append(grid);
  return disclosure;
}

function renderObjectAssumptionControls(item, kind) {
  const section = createElement("section", "object-assumptions");
  const fields = createElement("div", "object-assumption-fields");
  const hasRtsConstructionGroup = Boolean(getRtsBuildingGroup(item));
  if (hasRtsConstructionGroup) {
    const options = getRtsConstructionOptions(item);
    const constructionField = {
      ...JKSO_CONSTRUCTION_INPUT,
      options,
      labels: options.map((option) => JKSO_CONSTRUCTION_INPUT.labels[Number(option) - 1])
    };
    fields.append(createInputControl(constructionField,
      getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code), true,
      `${item.id}-${kind}`, "itemEstimate", item.key));
  }
  if (item.id === "VYB_05B") {
    fields.append(createInputControl(HALL_COVERAGE_INPUT,
      getItemEstimateValue(item, HALL_COVERAGE_INPUT.code), true,
      `${item.id}-${kind}`, "itemEstimate", item.key));
  }
  if (!fields.childElementCount) return null;
  section.append(fields);
  if (hasRtsConstructionGroup) {
    section.append(createElement("p", "project-estimate-input-note object-assumption-note",
      "Nabídka obsahuje jen konstrukční charakteristiky s cenou v odpovídající skupině RTS 2026; zvolená sazba v Kč/m³ OP se automaticky promítne do výpočtu. Zdroj: RTS 2026, s. 1, JKSO 801."));
  }
  return section;
}

function renderProjectEstimateMetrics(item, kind) {
  const estimate = getProjectPlanningEstimate(item, kind);
  if (!estimate) return null;
  const grid = createElement("div", "item-results-grid");
  const addMetric = (label, value, unit) => {
    const metric = createElement("div", "item-result-metric");
    metric.append(createElement("span", "", label));
    metric.append(createElement("strong", "", value === null
      ? "Neuvedeno" : formatValue(value, unit)));
    grid.append(metric);
  };
  if (estimate.type === "targeted") {
    const visibleLabels = TARGETED_RESULT_METRIC_LABELS[item.id] || [];
    estimate.metrics
      .filter(({ label }) => visibleLabels.some((prefix) => label.startsWith(prefix)))
      .forEach(({ label, value, unit }) => addMetric(label, value, unit));
    return grid;
  }
  if (estimate.type === "infrastructure") {
    if (estimate.isHeatingBoiler) {
      addMetric("Předběžná roční potřeba tepla", estimate.annualHeatDemand, "kWh");
      addMetric("Návrhový tepelný výkon", estimate.designCapacity, "kW");
      addMetric(estimate.pointLabel, estimate.pointCount, "ks");
      addMetric("Roční elektřina základního bydlení a vybraných vybaveností", estimate.annualElectricity, "kWh");
      addMetric("Špičkový elektrický příkon", estimate.peakElectricalLoad, "kW");
      if (estimate.cost === null) addMetric("Investiční cena", null, "");
      else addMetric("Modelový odhad investice", estimate.cost, "Kč");
      return grid;
    }
    const quantityLabel = item.id === "VYB_20A" ? "Objem retenční nádrže"
      : item.id === "VYB_20B_1" ? "Návrhový průtok vodovodu"
        : item.id === "VYB_20B_2" ? "Návrhový průtok splašků" : "Návrhový elektrický příkon";
    addMetric(quantityLabel, estimate.quantity, estimate.unit);
    if (item.id === "VYB_20B_3") addMetric("Roční elektřina základního bydlení a vybraných vybaveností", estimate.annualElectricity, "kWh");
    if (item.id !== "VYB_20A") addMetric("Délka nové trasy", estimate.networkLength, "m");
    if (estimate.pointCount > 0) addMetric(estimate.pointLabel, estimate.pointCount, "ks");
    addMetric("Odhad ceny realizace mimo RTS", estimate.cost, "Kč");
    return grid;
  }
  if (estimate.type === "polyclinic") {
    addMetric("Užitná plocha", estimate.usableArea, "m²");
    addMetric("Celková podlahová plocha", estimate.totalFloorArea, "m²");
    addMetric("Obestavěný prostor", estimate.volume, "m³");
    addMetric("Plocha areálu", estimate.siteArea, "m²");
    addMetric("Zpevněná plocha", estimate.pavedArea, "m²");
    addMetric("Investice včetně přirážky", estimate.cost, "Kč");
    return grid;
  }
  if (estimate.type === "school") return null;
  if (estimate.type === "item") {
    if (item.id === "VYB_10B_1") {
      addMetric("Počet modelových sálů",
        `${formatNumber(estimate.quantity)} ${estimate.quantity === 1 ? "sál" : "sálů"}`, "");
      addMetric("Užitná plocha", estimate.usableArea, "m²");
      addMetric("Celková podlahová plocha", estimate.totalFloorArea, "m²");
      addMetric("Obestavěný prostor", estimate.volume, "m³");
      addMetric("Plocha areálu", estimate.siteArea, "m²");
      addMetric("Zpevněná plocha", estimate.pavedArea, "m²");
      addMetric(estimate.rateSource === "RTS"
        ? "Investice podle RTS 2026" : "Odhad investice mimo RTS", estimate.cost, "Kč");
      return grid;
    }
    const scale = getStandardMetricScale(item, kind);
    if (estimate.usableArea !== null
      && !sameMetricValue(estimate.usableArea, getStandardMetricValue(item.usableArea, scale))) {
      addMetric("Projektová užitná plocha", estimate.usableArea, "m²");
    }
    if (estimate.totalFloorArea !== null && !sameMetricValue(estimate.totalFloorArea, estimate.usableArea)) {
      addMetric("Projektová celková podlahová plocha", estimate.totalFloorArea, "m²");
    }
    if (estimate.volume !== null
      && !sameMetricValue(estimate.volume, getStandardMetricValue(item.volume, scale, "m³"))) {
      addMetric("Projektový obestavěný prostor", estimate.volume, "m³");
    }
    if (estimate.siteArea !== null
      && !sameMetricValue(estimate.siteArea, getStandardMetricValue(item.parcelArea, scale, "m²"))) {
      addMetric("Projektová plocha pozemku", estimate.siteArea, "m²");
    }
    if (estimate.pavedArea !== null
      && !sameMetricValue(estimate.pavedArea, getStandardMetricValue(item.pavedArea, scale, "m²"))) {
      addMetric("Projektová zpevněná plocha", estimate.pavedArea, "m²");
    }
    if (estimate.cost !== null
      && !sameMetricValue(estimate.cost, getDisplayedCost(item, adjustedResult(item), kind))) {
      addMetric("Projektová investice", estimate.cost, "Kč");
    }
    return grid.childElementCount ? grid : null;
  }

  const scale = getStandardMetricScale(item, kind);
  const displayedUsableArea = getStandardMetricValue(item.usableArea, scale, "m²");
  const displayedVolume = getStandardMetricValue(item.volume, scale, "m³");
  const displayedPavedArea = getStandardMetricValue(item.pavedArea, scale, "m²");
  const displayedCost = getDisplayedCost(item, adjustedResult(item), kind);
  if (estimate.usableArea !== displayedUsableArea) {
    addMetric("Projektová užitná plocha", estimate.usableArea, "m²");
  }
  if (estimate.totalFloorArea !== estimate.usableArea) {
    addMetric("Projektová celková podlahová plocha", estimate.totalFloorArea, "m²");
  }
  if (estimate.volume !== displayedVolume) {
    addMetric("Projektový obestavěný prostor", estimate.volume, "m³");
  }
  addMetric("Projektová plocha areálu", estimate.siteArea, "m²");
  if (estimate.pavedArea !== displayedPavedArea) {
    addMetric("Projektová zpevněná plocha", estimate.pavedArea, "m²");
  }
  if (estimate.cost !== displayedCost) {
    addMetric("Projektová investice", estimate.cost, "Kč");
  }
  return grid;
}

function renderTargetedVerificationMetrics(estimate) {
  const grid = createElement("div", "item-results-grid targeted-verification-metrics");
  const verificationPattern = /ověřen|skutečná|skutečný|dostupnost|dojezd|započten|pokryt|vzdálenost|volná kapacita|bez doloženého|kontrola dostupnosti|výsledek|poptávka.*(?:nepokryt|odečtu)/i;
  for (const { label, value, unit } of estimate.metrics) {
    if (!verificationPattern.test(label)) continue;
    const metric = createElement("div", "item-result-metric");
    metric.append(createElement("span", "", label));
    metric.append(createElement("strong", "", value === null
      ? "Neověřeno" : formatValue(value, unit)));
    grid.append(metric);
  }
  return grid;
}

function renderTargetedVerificationMethods(estimate, item) {
  const equations = estimate?.verificationEquations ?? [];
  if (!equations.length) return null;
  const disclosure = createElement("details", "item-section targeted-verification-method");
  addDisclosureSummary(disclosure, "Postup ověření a podklady");
  const content = createElement("div", "calculation-content");
  equations.forEach((equation) => appendCalculationText(content, equation, item));
  disclosure.append(content);
  return disclosure;
}

function renderTargetedVerificationSection(item, kind) {
  const estimate = getTargetedPlanEstimate(item);
  if (!estimate) return null;
  const section = createElement("div", "targeted-verification");
  const controls = renderProjectEstimateControls(item, kind, estimate, "verification");
  if (controls) section.append(controls);
  section.append(renderTargetedVerificationMetrics(estimate));
  const methods = renderTargetedVerificationMethods(estimate, item);
  if (methods) section.append(methods);
  return section;
}

function appendCalculationLine(parent, label, parts, result, color) {
  const line = createElement("p", "calculation-equation calculation-line");
  line.append(createElement("strong", `estimate-value estimate-color-${color}`, label));
  if (parts.length) line.append(document.createTextNode(" = "));
  for (const part of parts) {
    if (typeof part === "string") {
      line.append(document.createTextNode(part));
      continue;
    }
    if (part.label) {
      line.append(createElement("span",
        `calculation-parameter-label estimate-value estimate-color-${part.color}`, `${part.label}: `));
    }
    const value = createElement("span", `estimate-value estimate-color-${part.color}`, part.value);
    if (part.label) value.title = part.label;
    line.append(value);
  }
  line.append(document.createTextNode(" = "),
    createElement("strong", `estimate-value estimate-color-${color}`, result));
  parent.append(line);
}

function getCalculationParameterValue(item, field, targetedCodes) {
  if (targetedCodes.has(field.code)) {
    return field.code === "MODEL_COST_PER_M3"
      ? getTargetedPlanCostRate(item) : getTargetedPlanInput(item, field.code);
  }
  if (field.code.startsWith("K_") || field.code.startsWith("POLY_")) {
    return getProjectEstimateValue(item, field.code);
  }
  if (PRIMARY_INPUTS.some((candidate) => candidate.code === field.code)) return getPrimary(field.code);
  if (SECONDARY_INPUTS.some((candidate) => candidate.code === field.code)
    || field.code === COVERAGE_OVERLAP_INPUT.code) return getSecondary(field.code);
  if (field.code.startsWith("BYTY_")) return getSecondary(field.code);
  return getItemEstimateValue(item, field.code);
}

function getCalculationParameterCandidates(item, kind) {
  const estimateType = getProjectPlanningEstimate(item, kind)?.type;
  const targetedFields = TARGETED_PLAN_INPUTS[item.id] || [];
  const targetedCodes = new Set(targetedFields.map((field) => field.code));
  const isHousingScenario = item.id === "VYB_07B" || RESIDENT_PARKING_IDS.has(item.id);
  const primaryFields = PRIMARY_INPUTS.filter((field) =>
    isHousingScenario || ["O_celk", "O_sidlo", "O_spad", "BJ_POCET_MANUAL",
      "PARKING_TOTAL_HPP", "S_uzemi", "S_zastav"].includes(field.code));
  const itemFields = estimateType === "school"
    ? ITEM_ESTIMATE_INPUTS.filter((field) => [
      "ITEM_VOLUME_PER_AREA", "ITEM_RTS_RATE", "ITEM_RTS_RATE_BASIS",
      "ITEM_RTS_FIXED_COST", "ITEM_COST_SURCHARGE"
    ].includes(field.code))
    : estimateType === "polyclinic"
      ? ITEM_ESTIMATE_INPUTS.filter((field) => field.code === "ITEM_RTS_FIXED_COST")
      : estimateType === "infrastructure" ? [] : ITEM_ESTIMATE_INPUTS;
  const fields = estimateType === "targeted"
    ? [...primaryFields, ...SECONDARY_INPUTS, COVERAGE_OVERLAP_INPUT, ...targetedFields]
    : [
      ...primaryFields,
      ...SECONDARY_INPUTS,
      COVERAGE_OVERLAP_INPUT,
      ...itemFields,
      ...(estimateType === "school" || estimateType === "polyclinic" ? PROJECT_ESTIMATE_INPUTS : []),
      ...(estimateType === "infrastructure" ? ALTERNATIVE_ESTIMATE_INPUTS : []),
      ...(item.id === "VYB_07B" ? HOUSING_CONTRIBUTION_INPUTS : [])
    ];

  return fields
    .filter((field) => ["number", "percent"].includes(field.kind))
    .map((field) => {
      const inputValue = numericOrNull(getCalculationParameterValue(item, field, targetedCodes));
      if (inputValue === null) return null;
      const color = field.color || ESTIMATE_FIELD_COLORS[field.code]
        || TARGETED_FIELD_COLORS[field.code] || getEstimateColorForInput(field) || "area";
      const values = field.kind === "percent"
        ? [formatNumber(inputValue * 100), formatNumber(inputValue)]
        : [formatNumber(inputValue)];
      const label = field.label || state.primaryLabels[field.code]
        || inputDescriptor(field.code)[INPUT_HEADERS.name] || field.code;
      return values.map((text) => ({ text, color, label }));
    })
    .filter(Boolean)
    .flat()
    .filter((candidate) => candidate.text)
    .sort((left, right) => right.text.length - left.text.length);
}

function appendCalculationParameterText(parent, value, item, kind, showLabels = true) {
  const candidates = getCalculationParameterCandidates(item, kind);
  if (!candidates.length) {
    parent.append(document.createTextNode(value));
    return;
  }
  const pattern = new RegExp(`(^|[^\\d.,])(${candidates
    .map((candidate) => candidate.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|")})(?=$|[^\\d.,])`, "g");
  let position = 0;
  for (const match of value.matchAll(pattern)) {
    const index = match.index ?? 0;
    parent.append(document.createTextNode(value.slice(position, index + match[1].length)));
    const candidate = candidates.find((entry) => entry.text === match[2]);
    if (showLabels && candidate?.label) {
      parent.append(createElement("span",
        `calculation-parameter-label estimate-value estimate-color-${candidate.color}`, `${candidate.label}: `));
    }
    const token = createElement("span", `estimate-value estimate-color-${candidate?.color || "area"}`, match[2]);
    if (candidate) token.title = candidate.label;
    parent.append(token);
    position = index + match[0].length;
  }
  parent.append(document.createTextNode(value.slice(position)));
}

function appendCalculationText(parent, equation, item = null, kind = "recommended") {
  const content = String(equation).replace(/\.$/, "");
  const firstSeparator = content.indexOf(" = ");
  const lastSeparator = content.lastIndexOf(" = ");
  if (firstSeparator < 0) {
    const line = createElement("p", "calculation-equation calculation-line");
    if (item) appendCalculationParameterText(line, content, item, kind, false);
    else line.textContent = content;
    parent.append(line);
    return;
  }
  const label = content.slice(0, firstSeparator);
  const expression = lastSeparator > firstSeparator
    ? content.slice(firstSeparator + 3, lastSeparator) : null;
  const result = expression === null ? content.slice(firstSeparator + 3) : content.slice(lastSeparator + 3);
  const normalizedLabel = label.toLocaleLowerCase("cs");
  const color = normalizedLabel.includes("zpevně")
    ? "paved"
    : normalizedLabel.includes("pozem") || normalizedLabel.includes("areál") ? "site"
      : normalizedLabel.includes("obestav") || normalizedLabel.includes("kubatur") ? "volume"
        : normalizedLabel.includes("celkov") && normalizedLabel.includes("ploch") ? "floor"
          : normalizedLabel.includes("invest") || normalizedLabel.includes("cena")
            || normalizedLabel.includes("náklad") || normalizedLabel.includes("zrn") ? "rts"
              : normalizedLabel.includes("navýšení") || normalizedLabel.includes("rezerv") ? "surcharge"
                : "area";
  const line = createElement("p", "calculation-equation calculation-line");
  if (item) {
    appendCalculationParameterText(line, label, item, kind);
    if (expression !== null) {
      line.append(document.createTextNode(" = "));
      appendCalculationParameterText(line, expression, item, kind);
    }
    line.append(document.createTextNode(" = "));
    appendCalculationParameterText(line, result, item, kind);
  } else {
    line.append(createElement("strong", `estimate-value estimate-color-${color}`, label));
    if (expression !== null) line.append(document.createTextNode(` = ${expression}`));
    line.append(document.createTextNode(" = "),
      createElement("strong", `estimate-value estimate-color-${color}`, result));
  }
  parent.append(line);
}

function renderProjectEstimateCalculation(item, kind, openingEquations = [], openingSources = []) {
  const estimate = getProjectPlanningEstimate(item, kind);
  const section = createElement("section", "calculation-numeric project-estimate-calculation");
  section.append(createElement("span", "calculation-kicker",
    estimate?.type === "targeted" ? "NÁVRHOVÝ VÝPOČET KAPACITY A NÁKLADŮ" : "PLOŠNÝ A INVESTIČNÍ VÝPOČET"));
  const controls = estimate ? renderProjectEstimateControls(item, kind, estimate) : null;
  if (controls) section.append(controls);
  openingEquations.forEach((equation) => appendCalculationText(section, equation, item, kind));
  const sources = [...openingSources];
  if (!estimate) {
    if (openingEquations.length) {
      section.append(createElement("p", "calculation-source-note",
        `Zdroje výchozích údajů: ${[...new Set(sources)].join(" ") || "hodnota převzatá z pracovního sešitu."}`));
    } else {
      section.append(createElement("p", "calculation-equation",
        "Plošný ani investiční propočet pro tuto položku není automaticky dostupný."));
    }
    return section;
  }
  const addLine = (label, parts, result, color) =>
    appendCalculationLine(section, label, parts, result, color);
  const addMissingLine = (label, explanation, color) => {
    const line = createElement("p", "calculation-equation calculation-line");
    line.append(createElement("strong", `estimate-value estimate-color-${color}`, label),
      document.createTextNode(`: ${explanation}`));
    section.append(line);
  };
  const addSpatialLines = (values, getColor) => {
    if (values.totalFloorArea !== null) {
      addLine("Celková podlahová plocha",
        [`${formatNumber(values.usableArea)} m² × `, {
          value: formatNumber(values.totalFloorFactor), color: getColor("floor"),
          label: "Koeficient celkové podlahové plochy / užitné plochy"
        }],
        `${formatNumber(values.totalFloorArea)} m²`, "floor");
    } else addMissingLine("Celková podlahová plocha", "chybí užitná plocha nebo její koeficient.", "floor");
    if (values.volume !== null) {
      addLine("Obestavěný prostor",
        [`${formatNumber(values.totalFloorArea)} m² × `, {
          value: formatNumber(values.volumeFactor), color: getColor("volume"),
          label: "Koeficient obestavěného prostoru"
        }, " m³/m²"],
        `${formatNumber(values.volume)} m³`, "volume");
    } else addMissingLine("Obestavěný prostor", "chybí podlahová plocha nebo koeficient obestavěného prostoru.", "volume");
    if (values.siteArea !== null) {
      addLine("Plocha areálu",
        [`${formatNumber(values.totalFloorArea)} m² × `, {
          value: formatNumber(values.siteAreaFactor), color: getColor("site"),
          label: "Koeficient plochy areálu / celkové podlahové plochy"
        }, " m²/m²"],
        `${formatNumber(values.siteArea)} m²`, "site");
    } else addMissingLine("Plocha areálu", "chybí podkladový koeficient plochy areálu.", "site");
    if (values.footprintArea !== null) {
      addLine("Půdorys objektu",
        [`${formatNumber(values.totalFloorArea)} m² × `, {
          value: formatNumber(values.footprintFactor), color: getColor("site"),
          label: "Podíl zastavěné plochy / celkové podlahové plochy"
        }, " m²/m²"],
        `${formatNumber(values.footprintArea)} m²`, "site");
    } else if (values.footprintFactor !== undefined) {
      addMissingLine("Půdorys objektu", "chybí podlahová plocha nebo koeficient půdorysu.", "site");
    }
    if (values.pavedArea !== null) {
      addLine("Zpevněná plocha",
        [`(${formatNumber(values.siteArea)} − ${formatNumber(values.footprintArea)}) m² × `,
        {
          value: formatNumber(values.pavedShare * 100), color: getColor("paved"),
          label: "Podíl zpevněné plochy ze zbytku areálu"
        }, " %"],
        `${formatNumber(values.pavedArea)} m²`, "paved");
    } else addMissingLine("Zpevněná plocha", "chybí podkladový koeficient zpevnění.", "paved");
  };

  if (estimate.type === "targeted") {
    estimate.equations.forEach((equation) => appendCalculationText(section, equation, item, kind));
    sources.push(...estimate.sources);
    section.append(createElement("p", "calculation-source-note", `Metodika a omezení: ${[...new Set(sources)].join(" ")}`));
    return section;
  }

  if (estimate.type === "public-space") {
    addLine("Kvalifikovaná zastavitelná plocha",
      [`${formatNumber(estimate.qualifyingAreaHa)} ha`],
      `${formatNumber(estimate.qualifyingAreaHa)} ha`, "area");
    addLine("Zákonné minimum veřejného prostranství",
      [`${formatNumber(estimate.qualifyingAreaHa)} ha × 500 m²/ha`],
      `${formatNumber(estimate.minimumArea)} m²`, "area");
    sources.push("§ 7 odst. 2 vyhlášky č. 501/2006 Sb.: nejméně 1 000 m² veřejného prostranství na každé 2 ha související zastavitelné plochy bydlení, rekreace, občanského vybavení nebo smíšené obytné. Pozemní komunikace se do výměry nezapočítávají.");
    section.append(createElement("p", "calculation-source-note", `Zdroj a podmínka použití: ${[...new Set(sources)].join(" ")}`));
    return section;
  }

  if (estimate.type === "item" && item.id === "VYB_10B_1") {
    const quantity = estimate.quantity;
    addLine("Počet modelových sálů", [], `${formatNumber(quantity)} ${quantity === 1 ? "sál" : "sálů"}`, "area");
    addLine("Užitná plocha",
      [`${formatNumber(quantity)} ks × ${formatNumber(estimate.usableAreaPerUnit)} m²/ks`],
      `${formatNumber(estimate.usableArea)} m²`, "area");
    addLine("Celková podlahová plocha",
      [`${formatNumber(estimate.usableArea)} m² × ${formatNumber(estimate.totalFloorFactor)} m²/m²`],
      `${formatNumber(estimate.totalFloorArea)} m²`, "floor");
    addLine("Obestavěný prostor",
      [`${formatNumber(estimate.totalFloorArea)} m² × ${formatNumber(estimate.volumePerArea)} m³/m²`],
      `${formatNumber(estimate.volume)} m³`, "volume");
    addLine("Plocha areálu",
      [`${formatNumber(quantity)} ks × ${formatNumber(estimate.siteAreaPerUnit)} m²/ks`],
      `${formatNumber(estimate.siteArea)} m²`, "site");
    addLine("Zpevněná plocha",
      [`${formatNumber(quantity)} ks × ${formatNumber(estimate.pavedAreaPerUnit)} m²/ks`],
      `${formatNumber(estimate.pavedArea)} m²`, "paved");
    if (estimate.zrn !== null) {
      addLine("Základní náklady (ZRN)",
        [`${formatNumber(estimate.volume)} m³ × ${formatCost(estimate.rate)}/m³ + ${formatCost(estimate.fixedCost)} fixních nákladů`],
        formatCost(estimate.zrn), "rts");
      addLine("Odhad investice",
        [`${formatCost(estimate.zrn)} × (1 + ${formatNumber(estimate.surcharge * 100)} %)`],
        formatCost(estimate.cost), "surcharge");
    } else {
      addMissingLine("Základní náklady (ZRN)",
        "RTS pro zvolenou konstrukční charakteristiku cenu neuvádí; sazbu lze zadat v předpokladech objektu.", "rts");
    }
    if (estimate.rateSource === "RTS") {
      sources.push(`${RTS_801_SOURCE} JKSO 801.4, konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}.`);
    } else if (estimate.rateSource === "user-estimate") {
      sources.push("Objemová sazba byla uživatelsky upravena; nejde o výchozí sazbu RTS.");
    } else {
      sources.push("Cena není pro zvolenou kombinaci v tabulce RTS uvedena. Plošné koeficienty a případná ručně zadaná sazba jsou pracovní předpoklady, nikoli normativní velikost nebo cenová nabídka.");
    }
    sources.push("Doporučení prověřit existenci a kapacitu sálu při více než 1 000 obyvatelích uvádí pravidlo CULT-MET-003.");
    section.append(createElement("p", "calculation-source-note", `Zdroj a povaha odhadu: ${[...new Set(sources)].join(" ")}`));
    return section;
  }

  if (estimate.type === "public-space") {
    addLine("Kvalifikovaná zastavitelná plocha",
      [`${formatNumber(estimate.qualifyingAreaHa)} ha`],
      `${formatNumber(estimate.qualifyingAreaHa)} ha`, "area");
    addLine("Zákonné minimum veřejného prostranství",
      [`${formatNumber(estimate.qualifyingAreaHa)} ha × 500 m²/ha`],
      `${formatNumber(estimate.minimumArea)} m²`, "area");
    sources.push("Výpočet podle § 7 odst. 2 vyhlášky č. 501/2006 Sb.: nejméně 1 000 m² veřejného prostranství na každé 2 ha související zastavitelné plochy bydlení, rekreace, občanského vybavení nebo smíšené obytné. Pozemní komunikace se do výměry nezapočítávají.");
    section.append(createElement("p", "calculation-source-note", `Zdroj a podmínka použití: ${[...new Set(sources)].join(" ")}`));
    return section;
  }

  if (estimate.type === "infrastructure") {
    const surface = getAggregateProjectSurfaceMetrics();
    if (estimate.isHeatingBoiler) {
      addLine("Roční elektřina základního bydlení a vybraných vybaveností",
        [`${formatNumber(estimate.modeledAnnualElectricity)} kWh/rok z tabulky + ${formatNumber(estimate.fallbackAnnualElectricity)} kWh/rok odhad pro nepokrytá data`],
        `${formatNumber(estimate.annualElectricity)} kWh/rok`, "volume");
      addLine("Modelová neelektrická roční energie",
        [`${formatNumber(estimate.modeledNonElectricEnergy)} kWh/rok; ${formatNumber(estimate.modeledEquipmentCount)} modelových položek včetně základního bydlení s energetickým údajem`],
        `${formatNumber(estimate.modeledNonElectricEnergy)} kWh/rok`, "area");
      addLine("Náhradní roční potřeba tepla pro chybějící údaje",
        [`${formatNumber(estimate.fallbackHeatingArea)} m² vytápěných ploch × ${formatNumber(estimate.heatDensity)} kWh/m²/rok`],
        `${formatNumber(estimate.fallbackHeatDemand)} kWh/rok`, "volume");
      addLine("Předběžná potřeba tepla pro ústřední vytápění",
        [`(${formatNumber(estimate.modeledNonElectricEnergy)} + ${formatNumber(estimate.fallbackHeatDemand)}) kWh/rok × ${formatNumber(estimate.heatingShare * 100)} %`],
        `${formatNumber(estimate.annualHeatDemand)} kWh/rok`, "area");
      addLine("Návrhový tepelný výkon kotelny",
        [`${formatNumber(estimate.annualHeatDemand)} kWh/rok ÷ ${formatNumber(estimate.fullLoadHours)} h/rok × ${formatNumber(estimate.reserveFactor)}`],
        `${formatNumber(estimate.designCapacity)} kW`, "volume");
      addLine("Počet kotlových modulů",
        [`zaokrouhlení nahoru (${formatNumber(estimate.designCapacity)} kW ÷ ${formatNumber(estimate.unitCapacity)} kW/modul)`],
        `${formatNumber(estimate.pointCount)} ks`, "fixed");
      if (estimate.pointRate === null) {
        addMissingLine("Investiční cena",
          "Chybí jednotkový odhad ceny kotlového modulu.", "rts");
      } else {
        addLine("Modelový odhad ceny kotlových modulů",
          [`${formatNumber(estimate.pointCount)} ks × ${formatCost(estimate.pointRate)}/ks`],
          formatCost(estimate.zrn), "rts");
        addLine("Modelový odhad včetně VRN a rezervy",
          [`${formatCost(estimate.zrn)} × (1 + ${formatNumber(estimate.costSurcharge * 100)} %)`],
          formatCost(estimate.cost), "surcharge");
      }
      sources.push("Energetické vstupy vycházejí z listu Model_Vybavenosti: roční celková energie = užitná plocha × měrná energie × koeficient PENB × geometrický koeficient; elektřina = celková energie × elektrický podíl; špičkový příkon = roční elektřina ÷ 8 760 × poměr špičky.");
      sources.push("Rozdíl celkové a elektrické energie je pouze modelový neelektrický energetický proxy, nikoli ověřená potřeba tepla. Pro položky bez energetických hodnot používá výpočet editovatelnou náhradní měrnou potřebu tepla a odhad elektrické zátěže; zdrojová tabulka ani cenový podklad kotelny nejsou úplné.");
      sources.push(`Ekvivalentní hodiny plného zatížení, podíl vytápění, výkonová rezerva, modulová kapacita a cena ${formatNumber(estimate.pointRate)} Kč/modul (${formatNumber(estimate.pointRate / estimate.unitCapacity)} Kč/kW) jsou upravitelné scénářové předpoklady. Jednotková cena je odvozena z výkonu modulu sazbou 10 000 Kč/kW, není doložena RTS ani nabídkou dodavatele a nezahrnuje kompletní technologii či stavební část. Návrhový výkon a počet modulů je nutné nahradit výpočtem podle energetického posudku a projektu.`);
      section.append(createElement("p", "calculation-source-note", `Zdroj a omezení: ${[...new Set(sources)].join(" ")}`));
      return section;
    } else if (item.id === "VYB_20A") {
      addLine("Odtoková plocha",
        [`${formatNumber(surface.roofArea)} m² střech + ${formatNumber(surface.pavedArea)} m² zpevněných ploch`],
        `${formatNumber(estimate.runoffArea)} m²`, "paved");
      addLine("Záchytný objem srážkové vody",
        [`${formatNumber(estimate.runoffArea)} m² × ${formatNumber(estimate.rainfallDepth)} mm ÷ 1 000 × ${formatNumber(estimate.runoffCoefficient)}`],
        `${formatNumber(estimate.quantity)} m³`, "area");
      sources.push("Modelová srážka 15 mm odpovídá pracovnímu údaji 150 m³/ha v listu TECH-HYDRO-001; součinitel odtoku je upravitelný předpoklad.");
      sources.push("Střešní plocha základního bytového fondu se odvozuje z HPP_byty a průměrné podlažnosti; další střechy a zpevněné plochy se přičítají z realizovaných vybaveností.");
      sources.push("Sazba 13 600 Kč/m³ je cenový proxy z TECH-COST-001, nikoli ověřená tržní cena konkrétní nádrže. Obsah výkopů, konstrukce a napojení ověřit projektem.");
    } else if (item.id === "VYB_20B_1" || item.id === "VYB_20B_2") {
      addLine("Průměrná denní potřeba vody",
        [`${formatNumber(estimate.population)} obyvatel × ${formatNumber(estimate.waterUse)} l/osobu/den ÷ 1 000`],
        `${formatNumber(estimate.averageDailyWater)} m³/den`, "area");
      const designFlow = item.id === "VYB_20B_2"
        ? estimate.averageDailyWastewater : estimate.averageDailyWater;
      if (item.id === "VYB_20B_2") {
        addLine("Průměrný denní odtok splašků",
          [`${formatNumber(estimate.averageDailyWater)} m³/den × ${formatNumber(estimate.wastewaterReturn * 100)} %`],
          `${formatNumber(estimate.averageDailyWastewater)} m³/den`, "paved");
      }
      addLine("Návrhová kapacita sítě",
        [`${formatNumber(designFlow)} m³/den × ${formatNumber(estimate.peakFactor)}`],
        `${formatNumber(estimate.quantity)} m³/den`, "volume");
      sources.push("Základní bytová populace O_celk se započítává bez ohledu na volbu městského odkupu bytů; odkup nemění počet obyvatel ani potřebu vody a kanalizace.");
      const lengthLabel = item.id === "VYB_20B_1" ? "Délka nového vodovodního potrubí" : "Délka nové splaškové kanalizace";
      const pipeLabel = item.id === "VYB_20B_1" ? "Cena vodovodního potrubí včetně pokládky" : "Cena splaškového potrubí včetně pokládky";
      addLine(lengthLabel, [], `${formatNumber(estimate.networkLength)} m`, "site");
      addLine(pipeLabel,
        [`${formatNumber(estimate.networkLength)} m × ${formatCost(estimate.lineRate)}/m`],
        formatCost(estimate.lineCost), "rts");
      if (item.id === "VYB_20B_1") {
        addLine("Cena vodoměrných a armaturních šachet",
          [`${formatNumber(estimate.pointCount)} ks × ${formatCost(estimate.pointRate)}/ks`],
          formatCost(estimate.pointCost), "fixed");
      } else {
        const spacing = numericOrNull(getItemEstimateValue(item, "ALT_SEWER_MANHOLE_SPACING_M"));
        addLine("Počet revizních šachet",
          [`zaokrouhlení nahoru (${formatNumber(estimate.networkLength)} m ÷ ${formatNumber(spacing)} m) + 1`],
          `${formatNumber(estimate.pointCount)} ks`, "fixed");
        addLine("Cena revizních šachet",
          [`${formatNumber(estimate.pointCount)} ks × ${formatCost(estimate.pointRate)}/ks`],
          formatCost(estimate.pointCost), "fixed");
      }
      const totalAllocation = 4 * 420 * 13600 * 0.35;
      const defaultPointCost = item.id === "VYB_20B_1" ? 150000 : 31 * 25000;
      const derivedDefaultRate = (totalAllocation - defaultPointCost) / 1500;
      addLine("Původ výchozí sazby za metr (pracovní odvození)",
        [`(4 uzly × 420 m³/uzel × 13 600 Kč/m³ × 35 % − ${formatCost(defaultPointCost)} bodových prvků) ÷ 1 500 m`],
        `${formatCost(derivedDefaultRate)}/m`, "rts");
      sources.push("ČSÚ, „Vodovody, kanalizace a vodní toky – data za rok 2023“, tabulky 1.1.1 a 1.1.4: 325 312 tis. m³ domácnostem ÷ 10 279 174 zásobovaných obyvatel ÷ 365 × 1 000 = 86,7 l/os./den, zaokrouhleno na 87. Špičkový faktor a návratnost splašků jsou pracovní odhady.");
      sources.push("Výchozí cena potrubí za metr je odvozena z alokace 35 % scénáře TECH-COST-001 při trase 1 500 m po odečtení odhadované ceny šachet. Nejde o ověřenou tabulkovou cenu za metr; rozpočet závisí na dimenzi, materiálu, hloubce, povrchu, pažení, čerpání a podmínkách správce.");
    } else {
      addLine("Souhrnná podlahová plocha objektů", [], `${formatNumber(surface.totalFloorArea)} m²`, "floor");
      addLine("Špičkový příkon z energetické tabulky",
        [`součet údajů ${formatNumber(estimate.modeledPeakElectricalLoad)} kW u ${formatNumber(estimate.modeledEquipmentCount)} modelových položek včetně základního bydlení`],
        `${formatNumber(estimate.modeledPeakElectricalLoad)} kW`, "volume");
      addLine("Náhradní příkon pro položky bez údajů",
        [`${formatNumber(estimate.fallbackFloorArea)} m² × ${formatNumber(estimate.energyDensity)} W/m² ÷ 1 000`],
        `${formatNumber(estimate.fallbackPeakElectricalLoad)} kW`, "area");
      addLine("Celkový návrhový příkon",
        [`${formatNumber(estimate.modeledPeakElectricalLoad)} kW z tabulky + ${formatNumber(estimate.fallbackPeakElectricalLoad)} kW náhradní odhad`],
        `${formatNumber(estimate.quantity)} kW`, "volume");
      addLine("Roční elektřina základního bydlení a vybraných vybaveností",
        [`${formatNumber(estimate.modeledAnnualElectricity)} kWh/rok z tabulky + odhad podle poměru špičky ${formatNumber(estimate.peakToAverageRatio)}`],
        `${formatNumber(estimate.annualElectricity)} kWh/rok`, "volume");
      addLine("Délka nového elektrického kabelového vedení", [], `${formatNumber(estimate.networkLength)} m`, "site");
      addLine("Cena kabelového vedení včetně pokládky",
        [`${formatNumber(estimate.networkLength)} m × ${formatCost(estimate.lineRate)}/m`],
        formatCost(estimate.lineCost), "rts");
      addLine("Počet trafostanic",
        [`zaokrouhlení nahoru (${formatNumber(estimate.quantity)} kW ÷ ${formatNumber(estimate.transformerCapacity)} kW/stanici)`],
        `${formatNumber(estimate.pointCount)} ks`, "fixed");
      addLine("Cena trafostanic včetně osazení",
        [`${formatNumber(estimate.pointCount)} ks × ${formatCost(estimate.pointRate)}/ks`],
        formatCost(estimate.pointCost), "fixed");
      addLine("Výchozí cena elektrického kabelu za metr (pracovní odvození)",
        [`(4 uzly × 420 m³/uzel × 13 600 Kč/m³ × 30 % − 1 × 3 000 000 Kč za trafostanici) ÷ 1 500 m`],
        `${formatCost((4 * 420 * 13600 * 0.30 - 3000000) / 1500)}/m`, "rts");
      sources.push("Základní bytový fond podle O_celk se započítává vždy; další realizované vybavenosti se přičítají. Městský odkup bytů proto nemění počet ani energetickou potřebu bytů.");
      sources.push("Modelový roční odběr elektřiny a špičkový příkon se načítají z energetických sloupců bytového fondu a vybraných vybaveností v listu Model_Vybavenosti. Pro objekty bez kompletních energetických hodnot zůstává editovatelná náhrada podle plochy: výchozích 50 W/m² a poměr špičky 4.");
      sources.push("Kapacita a cena trafostanice i kabelové vedení jsou upravitelné odhady, nikoli normové hodnoty nebo RTS ceny.");
      sources.push("Výchozí cena kabelu za metr je odvozena z alokace 30 % scénáře TECH-COST-001 při trase 1 500 m po odečtení pracovního odhadu trafostanice za 3 000 000 Kč; dimenzi, napětí a podmínky připojení musí určit správce a projektant.");
    }
    const investmentLabel = item.id === "VYB_20A"
      ? "Odhad ceny retenční nádrže a stavebních prací"
      : "Odhad ceny potrubí/kabelu a bodových prvků";
    addLine(investmentLabel,
      item.id === "VYB_20A"
        ? [`${formatNumber(estimate.quantity)} m³ × ${formatCost(estimate.lineRate)}/m³`]
        : [`${formatCost(estimate.lineCost)} potrubí/kabel + ${formatCost(estimate.pointCost)} bodové prvky`],
      formatCost(estimate.zrn), "rts");
    addLine("Odhad investice včetně VRN a rezervy",
      [`${formatCost(estimate.zrn)} × (1 + ${formatNumber(estimate.costSurcharge * 100)} %)`],
      formatCost(estimate.cost), "surcharge");
  } else if (estimate.type === "polyclinic") {
    const currentPopulation = numericOrNull(getPrimary("O_sidlo"));
    const projectPopulation = numericOrNull(getPrimary("O_celk"));
    const catchmentPopulation = numericOrNull(getPrimary("O_spad"));
    addLine("Obyvatelé ve spádu po záměru",
      [`${formatNumber(currentPopulation)} + ${formatNumber(projectPopulation)} + ${formatNumber(catchmentPopulation)} obyvatel`],
      `${formatNumber(estimate.population)} obyvatel`, "population");
    addLine("Ekvivalent poliklinických center",
      [`${formatNumber(estimate.population)} obyvatel ÷ `,
        {
          value: formatNumber(estimate.peoplePerCentre), color: "area",
          label: "Obyvatel na jedno modelové poliklinické centrum"
        }, " obyvatel/centrum"],
      `${formatNumber(estimate.centerEquivalent, 3)} centra`, "area");
    addLine("Užitná plocha",
      [`${formatNumber(estimate.centerEquivalent, 3)} centra × `,
        {
          value: formatNumber(estimate.usableAreaPerCentre), color: "area",
          label: "Užitná plocha jednoho modelového poliklinického centra"
        }, " m²/centrum"],
      `${formatNumber(estimate.usableArea)} m²`, "area");
    sources.push("Počty obyvatel: uživatelské vstupy řešeného území a spádu.");
    if (estimate.costPerCubicMeter !== null) {
      sources.push(`${RTS_801_SOURCE} JKSO 801.1, konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}.`);
    } else {
      sources.push(`${RTS_801_SOURCE} Pro tento propočet není vybraná cenová sazba.`);
    }
    sources.push("Plocha centra a hranice počtu obyvatel: pracovní scénář zdravotnického clusteru, nikoli zákonný požadavek.");
    addSpatialLines({
      usableArea: estimate.usableArea,
      totalFloorFactor: estimate.totalFloorFactor,
      totalFloorArea: estimate.totalFloorArea,
      volumeFactor: estimate.volumeFactor,
      volume: estimate.volume,
      siteAreaFactor: estimate.siteAreaFactor,
      siteArea: estimate.siteArea,
      footprintFactor: estimate.footprintFactor,
      footprintArea: estimate.footprintArea,
      pavedShare: estimate.pavedShare,
      pavedArea: estimate.pavedArea
    }, (color) => color);
    if (estimate.zrn !== null) {
      addLine("Základní náklady (ZRN)",
        [`${formatNumber(estimate.volume)} m³ × `, {
          value: formatNumber(estimate.costPerCubicMeter), color: "rts",
          label: "ZRN polikliniky na 1 m³ obestavěného prostoru"
        },
          ` Kč/m³ + ${formatCost(estimate.fixedCost)} fixních nákladů`],
        formatCost(estimate.zrn), "rts");
      addLine("Odhad investice",
        [`${formatCost(estimate.zrn)} × (1 + `,
          {
            value: formatNumber(estimate.costSurcharge * 100), color: "surcharge",
            label: "Přirážka VRN a rezervy k ZRN"
          }, " %)"],
        formatCost(estimate.cost), "surcharge");
    } else {
      addMissingLine("Základní náklady (ZRN)",
        "RTS pro vybranou konstrukční charakteristiku neuvádí cenu; doplňte sazbu v projektových předpokladech.", "rts");
    }
  } else {
    if (estimate.type === "school") {
      const usableAreaSource = calculationSource("Užitná plocha", item.metricSources?.usableArea);
      const sourceVolumeSource = calculationSource("Obestavěný prostor", item.metricSources?.volume);
      const rtsSource = calculationSource("Sazba RTS a fixní náklady", item.sourceFormulas?.cost);
      if (usableAreaSource) sources.push(usableAreaSource);
      if (sourceVolumeSource) sources.push(sourceVolumeSource);
      if (getRtsBuildingGroup(item)) {
        sources.push(`${RTS_801_SOURCE} JKSO ${getRtsBuildingGroup(item)}, konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}.`);
      } else if (rtsSource) {
        sources.push(rtsSource);
      }
      addLine("Základní užitná plocha", [], `${formatNumber(estimate.baseUsableArea)} m²`, "area");
      if (estimate.upliftPercent) {
        addLine("Rezerva užitné plochy",
          [`${formatNumber(estimate.standardUsableArea)} m² × `,
            {
              value: formatNumber(estimate.upliftPercent), color: "surcharge",
              label: "Doporučené navýšení kapacity"
            }, " %"],
          `${formatNumber(estimate.reserveUsableArea)} m²`, "surcharge");
        addLine("Užitná plocha včetně rezervy",
          [`${formatNumber(estimate.standardUsableArea)} m² + ${formatNumber(estimate.reserveUsableArea)} m²`],
          `${formatNumber(estimate.usableArea)} m²`, "area");
      } else {
        addLine("Užitná plocha pro zvolený výsledek", [], `${formatNumber(estimate.usableArea)} m²`, "area");
      }
      addSpatialLines({
        usableArea: estimate.usableArea,
        totalFloorFactor: estimate.totalFloorFactor,
        totalFloorArea: estimate.totalFloorArea,
        volumeFactor: estimate.volumeFactor,
        volume: estimate.volume,
        siteAreaFactor: estimate.siteAreaFactor,
        siteArea: estimate.siteArea,
        footprintFactor: estimate.footprintFactor,
        footprintArea: estimate.footprintArea,
        pavedShare: estimate.pavedShare,
        pavedArea: estimate.pavedArea
      }, (color) => color);
      if (estimate.zrn !== null) {
        addLine("Základní náklady (ZRN)",
          [`${formatNumber(estimate.costBasis)} ${estimate.costRateBasis === "volume" ? "m³" : text(item.unit, "jednotek")} × `,
            { value: formatNumber(estimate.costRate), color: "rts", label: "Cena dle měrné jednotky" },
            ` Kč/${estimate.costRateBasis === "volume" ? "m³" : text(item.unit, "jednotka")} + ${formatCost(estimate.fixedCost)} fixních nákladů`],
          formatCost(estimate.zrn), "rts");
        addLine("Odhad investice",
          [`${formatCost(estimate.zrn)} × (1 + `,
            {
              value: formatNumber(estimate.costSurcharge * 100), color: "surcharge",
              label: "VRN a rezerva"
            }, " %)"],
          formatCost(estimate.cost), "surcharge");
      } else {
        addMissingLine("Odhad investice", "chybí zdrojová sazba RTS nebo její měrná jednotka.", "rts");
      }
    } else {
      const parkingAssumption = PARKING_SHARE_ASSUMPTIONS[item.id];
      if (parkingAssumption) {
        const population = numericOrNull(getPrimary("O_celk")) ?? 0;
        addLine("Počet jednotek v propočtu",
          [`zaokrouhlení nahoru (${formatNumber(population)} obyvatel × `,
            {
              value: formatNumber(getParkingShare(item)), color: "area",
              label: "Scénář poptávky: podíl obyvatel"
            }, " %)"],
          `${formatNumber(estimate.quantity, 2)} ${text(item.unit, "jednotek")}`, "area");
        sources.push("Podíl obyvatel pro odhad počtu jednotek je pracovní, uživatelsky upravitelný předpoklad; nejde o RTS sazbu ani zákonný požadavek.");
      } else {
        addLine("Počet jednotek v propočtu", [],
          `${formatNumber(estimate.quantity, 2)} ${text(item.unit, "jednotek")}`, "area");
        if (item.id === "VYB_05B") {
          sources.push("Volba „Samostatná nová hala požadována“ vytváří jeden modelový objekt. Výchozí předpoklady na halu: 1 800 m² užitné plochy, koeficient celkové plochy 1,1, 4,2 m³/m², areál 6 000 m² a zpevněná plocha 3 000 m². Jde o editovatelný scénář pro chybějící jednotky v sešitu, nikoli normový nebo projektový návrh.");
        }
      }
      if (estimate.usableArea !== null) {
        addLine("Užitná plocha",
          [`${formatNumber(estimate.quantity, 2)} ${text(item.unit, "jednotek")} × `,
            {
              value: formatNumber(estimate.usableAreaPerUnit), color: "area",
              label: "Užitná plocha na jednotku"
            }, " m²/jednotku"],
          `${formatNumber(estimate.usableArea)} m²`, "area");
        const source = calculationSource("Užitná plocha", item.metricSources?.usableArea);
        if (source) sources.push(source);
      } else addMissingLine("Užitná plocha", "v sešitu chybí podklad na jednotku; lze jej doplnit výše.", "area");
      if (estimate.totalFloorArea !== null) {
        addLine("Celková podlahová plocha",
          [`${formatNumber(estimate.usableArea)} m² × `,
            {
              value: formatNumber(estimate.totalFloorFactor), color: "floor",
              label: "Koeficient celkové podlahové plochy"
            }],
          `${formatNumber(estimate.totalFloorArea)} m²`, "floor");
      } else addMissingLine("Celková podlahová plocha", "chybí užitná plocha nebo koeficient.", "floor");
      if (estimate.volume !== null) {
        addLine("Obestavěný prostor",
          [`${formatNumber(estimate.totalFloorArea)} m² × `,
            {
              value: formatNumber(estimate.volumePerArea), color: "volume",
              label: "Koeficient obestavěného prostoru"
            }, " m³/m²"],
          `${formatNumber(estimate.volume)} m³`, "volume");
        const source = calculationSource("Obestavěný prostor", item.metricSources?.volume);
        if (source) sources.push(source);
      } else addMissingLine("Obestavěný prostor", "chybí koeficient obestavěného prostoru.", "volume");
      if (estimate.siteArea !== null) {
        addLine("Plocha pozemku",
          [`${formatNumber(estimate.quantity, 2)} × `,
            {
              value: formatNumber(estimate.siteAreaPerUnit), color: "site",
              label: "Plocha pozemku na jednotku"
            }, " m²/jednotku"],
          `${formatNumber(estimate.siteArea)} m²`, "site");
      } else addMissingLine("Plocha pozemku", "chybí podkladový koeficient; lze jej doplnit výše.", "site");
      if (estimate.pavedArea !== null) {
        addLine("Zpevněná plocha",
          [`${formatNumber(estimate.quantity, 2)} × `,
            {
              value: formatNumber(estimate.pavedAreaPerUnit), color: "paved",
              label: "Zpevněná plocha na jednotku"
            }, " m²/jednotku"],
          `${formatNumber(estimate.pavedArea)} m²`, "paved");
      } else addMissingLine("Zpevněná plocha", "chybí podkladový koeficient; lze jej doplnit výše.", "paved");
      if (estimate.zrn !== null) {
        const volumeBasis = estimate.rateBasis === "volume";
        addLine(estimate.rateSource === "RTS" ? "Základní náklady (ZRN)" : "Základní náklady (odhad mimo RTS)",
          [`${formatNumber(volumeBasis ? estimate.volume : estimate.quantity, 2)} ${volumeBasis ? "m³" : text(item.unit, "jednotek")} × `,
            { value: formatNumber(estimate.rate), color: "rts", label: "Cena dle měrné jednotky" },
            ` Kč/${volumeBasis ? "m³" : text(item.unit, "jednotka")} + ${formatCost(estimate.fixedCost)} fixních nákladů`],
          formatCost(estimate.zrn), "rts");
        if (estimate.rateSource === "RTS") {
          if (getRtsBuildingGroup(item)) {
            sources.push(`${RTS_801_SOURCE} JKSO ${getRtsBuildingGroup(item)}, konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}.`);
          } else {
            const source = calculationSource("Sazba RTS", item.sourceFormulas?.cost);
            if (source) sources.push(source);
          }
        } else if (estimate.rateSource === "user-estimate") {
          sources.push(`Cenový odhad mimo RTS: uživatelsky zadaná nebo výchozí modelová sazba ${formatCost(estimate.rate)}/${volumeBasis ? "m³" : text(item.unit, "jednotka")}; nejde o ověřenou sazbu RTS.`);
        } else if (estimate.rateSource === "workbook-unit-cost") {
          const source = calculationSource("Jednotková cena mimo RTS", item.sourceFormulas?.unitPrice);
          sources.push(source
            ? `Cenový odhad mimo RTS: ${source}`
            : "Jednotková cena je převzatá z pracovního modelu, nejde o ověřenou sazbu RTS.");
        } else if (estimate.rateSource === "workbook-total-cost-derived") {
          const source = calculationSource("Modelový náklad", item.sourceFormulas?.cost);
          sources.push(`Cenový odhad mimo RTS: jednotková sazba je odvozena z modelového nákladu dělením počtem jednotek a předpokládaným faktorem 1,23${source ? ` (${source})` : ""}.`);
        } else if (estimate.rateSource === "scenario-estimate" && item.id === "VYB_04B") {
          const sourceAddress = item.sourceFormulas?.unitPrice?.sourceAddress;
          sources.push(`RTS v dostupném podkladu cenu veřejné parkové zeleně neuvádí. Orientační sazba ${formatCost(estimate.rate)}/park, tj. ${formatNumber(estimate.rate / estimate.usableAreaPerUnit)} Kč/m² při ${formatNumber(estimate.usableAreaPerUnit)} m²/park, vychází z pracovního scénáře zeleně${sourceAddress ? ` (zdrojová buňka ${sourceAddress})` : ""}; nejde o ověřenou sazbu RTS ani nabídku dodavatele.`);
        } else if (estimate.rateSource === "scenario-estimate") {
          sources.push("Cenový odhad mimo RTS: orientační objemová sazba 13 645 Kč/m³ je převzata jako srovnávací předpoklad ze scénáře polikliniky v pracovním zdravotnickém clusteru; plošné a objemové koeficienty jsou pracovní odhady upravitelné u této položky.");
        } else if (estimate.rateSource === "comparable-estimate") {
          sources.push(`Cenový odhad mimo RTS: medián jednotkových cen položek stejné kategorie a měrné jednotky (${formatCost(estimate.rate)}/${volumeBasis ? "m³" : text(item.unit, "jednotka")}).`);
        }
        addLine("Odhad investice",
          [`${formatCost(estimate.zrn)} × (1 + `,
            {
              value: formatNumber(estimate.surcharge * 100), color: "surcharge",
              label: "VRN a rezerva"
            }, " %)"],
          formatCost(estimate.cost), "surcharge");
      } else addMissingLine("Investice", "není dostupná sazba ani srovnávací cenový podklad; doplňte pracovní odhad a jeho měrný základ.", "rts");
    }
  }
  if (sources.length) {
    section.append(createElement("p", "calculation-source-note", `Zdroje výchozích údajů: ${[...new Set(sources)].join(" ")}`));
  }
  return section;
}

function renderInfrastructureResults(item, kind) {
  const section = createElement("section", "item-results");
  section.append(createElement("h4", "", RESULTS_HEADING));
  const metrics = renderProjectEstimateMetrics(item, kind);
  if (metrics) section.append(metrics);
  return section;
}

function renderHousingContributionCalculation(item, kind) {
  const estimate = getHousingContributionEstimate(item);
  if (!estimate) return null;
  const section = createElement("section", "calculation-numeric housing-contribution-calculation");
  section.append(createElement("span", "calculation-kicker", "PLOŠNÝ A INVESTIČNÍ VÝPOČET"));
  const controls = renderProjectEstimateControls(item, kind, null);
  if (controls) section.append(controls);
  appendCalculationLine(section, "Byty v bytových domech v záměru",
    [`${formatNumber(estimate.totalUnits)} bytových jednotek a rodinných domů × `,
      {
        value: `${formatNumber(estimate.apartmentShare * 100)} %`, color: "area",
        label: "Podíl bytů v bytových domech"
      },
      ", zaokrouhleno na celé byty"],
    `${formatNumber(estimate.totalApartments)} bytů`, "area");
  appendCalculationLine(section, "Byty odkupované městem",
    [`${formatNumber(estimate.totalApartments)} bytů × ${formatNumber(estimate.buyoutShare * 100)} % odkupu, zaokrouhleno na celé byty`],
    `${formatNumber(estimate.municipalApartments)} bytů`, "area");
  appendCalculationLine(section, "Plocha všech bytů",
    [`${formatNumber(estimate.totalApartments)} bytů × `,
      {
        value: `${formatNumber(estimate.areaPerHome)} m²/byt`, color: "area",
        label: "Užitná plocha jednoho bytu"
      }],
    `${formatNumber(estimate.portfolioArea)} m²`, "area");
  appendCalculationLine(section, "Obestavěný prostor celého portfolia",
    [`zaokrouhlení (${formatNumber(estimate.portfolioArea)} m² × `,
      {
        value: `${formatNumber(estimate.volumePerArea)} m³/m²`, color: "volume",
        label: "Obestavěný prostor na 1 m² bytové plochy"
      }, ")"],
    `${formatNumber(estimate.portfolioVolume)} m³`, "volume");
  appendCalculationLine(section, "Modelová hodnota celého portfolia",
    [`${formatNumber(estimate.portfolioVolume)} m³ × `,
      {
        value: `${formatCost(estimate.valuePerCubicMeter)}/m³`, color: "rts",
        label: "Modelová hodnota na 1 m³ obestavěného prostoru"
      }],
    formatCost(estimate.portfolioValue), "rts");
  appendCalculationLine(section, "Plocha bytů odkupovaných městem",
    [`${formatNumber(estimate.municipalApartments)} bytů × `,
      {
        value: `${formatNumber(estimate.areaPerHome)} m²/byt`, color: "area",
        label: "Užitná plocha jednoho bytu"
      }],
    `${formatNumber(estimate.municipalArea)} m²`, "area");
  appendCalculationLine(section, "Obestavěný prostor městského podílu",
    [`zaokrouhlení (${formatNumber(estimate.municipalArea)} m² × `,
      {
        value: `${formatNumber(estimate.volumePerArea)} m³/m²`, color: "volume",
        label: "Obestavěný prostor na 1 m² bytové plochy"
      }, ")"],
    `${formatNumber(estimate.municipalVolume)} m³`, "volume");
  appendCalculationLine(section, "Modelová hodnota městského podílu",
    [`${formatNumber(estimate.municipalVolume)} m³ × `,
      {
        value: `${formatCost(estimate.valuePerCubicMeter)}/m³`, color: "rts",
        label: "Modelová hodnota na 1 m³ obestavěného prostoru"
      }],
    formatCost(estimate.municipalValue), "rts");
  appendCalculationLine(section, "Modelový náklad městských bytů (ušlá marže)",
    [`${formatCost(estimate.municipalValue)} × `,
      {
        value: `${formatNumber(estimate.lostMarginShare * 100)} %`, color: "surcharge",
        label: "Podíl ušlé marže"
      }],
    formatCost(estimate.lostMargin), "surcharge");
  section.append(createElement("p", "calculation-source-note",
    "Zdroje: počet osob na byt je výchozí koeficient scénáře; podíl odkupu a ušlá marže jsou uživatelské vstupy; 52 m²/byt, 3,9 m³/m² a sazba 12 400 Kč/m³ vycházejí z pracovního bytového scénáře. Jde o odhad nákladové báze a ušlé marže, nikoli o odhad tržní prodejní ceny."));
  return section;
}

function getDisplayedCost(item, calc, kind) {
  const housingEstimate = getHousingContributionEstimate(item);
  if (housingEstimate) return housingEstimate.lostMargin;
  if (isResidentParkingItem(item)) {
    const count = calc.required;
    const source = state.items.find((candidate) => candidate.id === RESIDENT_PARKING_ID);
    const modelCount = numericOrNull(source?.demand) ?? numericOrNull(source?.proposed);
    if (count === null || count === undefined || modelCount === null || modelCount <= 0
      || item.cost === null || item.cost < 0) return null;
    return item.cost * count / modelCount;
  }
  const isDelta = kind === "recommended" && getPrimary("REZIM_STD") !== "STANDARD"
    && hasStandardComparison(item) && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
  if (isDelta && getStandardComparison(item).delta === null) return null;
  if (hasStandardComparison(item) && (hasAppliedStandardTarget(item) || isDelta)) {
    const scale = getStandardMetricScale(item, kind, calc);
    if (item.cost !== null && item.cost >= 0 && Number.isFinite(scale)) return item.cost * scale;
  }
  return calc.cost;
}

function updateRequirementCardSummary(card, item, calc) {
  const quantity = card.querySelector(".requirement-quantity");
  if (quantity) {
    if (item.id === "VYB_05B") {
      const hallCount = getItemEstimateQuantity(item);
      quantity.textContent = hallCount > 0
        ? `${formatNumber(hallCount)} ${hallCount === 1 ? "hala" : "haly"}`
        : "bez samostatné haly";
    } else {
      const summaryValue = getSummaryRequirementQuantity(calc);
      quantity.textContent = summaryValue === null
        ? "bez počtu" : `${formatNumber(summaryValue, 2)} ${text(getRequirementUnit(item), "")}`;
    }
  }
  const capacitySummary = card.querySelector(".requirement-capacity");
  if (capacitySummary && calc.usesDemandBasedCapacity) {
    capacitySummary.textContent = getCapacitySummaryText(item, calc);
    capacitySummary.hidden = !calc.isCapacityVerified;
  }
}

function hasStandardComparison(item) {
  if (isResidentParkingItem(item)) return false;
  return item.legalRules.some((rule) => !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id])))
    && item.recommendedRules.length > 0;
}

function getStandardComparison(item) {
  const base = getStandardBaseQuantity(item);
  const uplift = getRecommendedIncrease(item);
  if (uplift === null) {
    return { legal: null, base, recommended: null, uplift: null, delta: null, cost: null, appliedCapacity: 0 };
  }
  const recommended = base === null ? null : base + uplift;
  const reserveCapacity = getAllowedCapacityUse(item, getVerifiedCapacity(item), uplift);
  const appliedCapacity = Math.min(reserveCapacity, uplift);
  const delta = Math.max(0, uplift - appliedCapacity);
  const cost = base > 0 && item.cost !== null ? item.cost * delta / base : null;
  return { legal: null, base, recommended, uplift, delta, cost, appliedCapacity };
}

function getCapacityUnit(item) {
  if (isHealthcareStaffingItem(item)) return "modelová souč. zátěž";
  return CAPACITY_DEMAND_UNITS[item.id] || item.unit;
}

function getCapacityResultUnit(item) {
  return DEMAND_LABELS[item.id] || getCapacityUnit(item);
}

function usesDemandBasedCapacity(item) {
  return Object.hasOwn(CAPACITY_DEMAND_UNITS, item.id);
}

function isHealthcareStaffingItem(item) {
  return ["VYB_17", "VYB_18", "VYB_19", "VYB_09A_4"].includes(item.id)
    && isHealthcareItem(item);
}

function isHealthcareItem(item) {
  return String(item.category ?? "").toLocaleUpperCase("cs").includes("ZDRAVOTNICTVÍ");
}

function getHealthcareStaffingBasis(item) {
  const fallback = HEALTHCARE_POPULATION_PER_FTE[item.id];
  if (!fallback) return null;
  const value = numericOrNull(state.healthcarePopulationPerFte?.[item.id]);
  return { ...fallback, value: value !== null && value > 0 ? value : fallback.value };
}

function isRequirementSatisfied(calc) {
  return calc.usesDemandBasedCapacity ? calc.remainingCapacity === 0 : calc.remaining === 0;
}

function modelCountUnit(item) {
  return CAPACITY_QUANTITY_IDS.has(item.id) || item.id === "VYB_04A"
    ? "zařízení/bodů" : item.unit;
}

function getDemandUnit(item) {
  if (DEMAND_LABELS[item.id]) return DEMAND_LABELS[item.id];
  if (isHealthcareStaffingItem(item)) return getCapacityUnit(item);
  if (usesDemandBasedCapacity(item) || CAPACITY_QUANTITY_IDS.has(item.id)) return item.unit;
  return "";
}

function getDemandLabel(item) {
  if (isHealthcareStaffingItem(item)) return "Odhadovaný přírůstek personální potřeby";
  return getDemandUnit(item) ? "Kapacitní poptávka" : "Kapacitní poptávka (jednotku ověřit)";
}

function getCapacityEdit(item) {
  return state.capacityEdits[item.key] || { quantity: "", source: "", note: "", date: "" };
}

function getParkingShare(item) {
  const saved = numericOrNull(state.parkingShares[item.key]);
  return saved !== null && saved >= 0 && saved <= 100
    ? saved : PARKING_SHARE_ASSUMPTIONS[item.id]?.defaultPercent ?? null;
}

function renderParkingDemand(item) {
  const config = PARKING_SHARE_ASSUMPTIONS[item.id];
  if (!config) return null;
  const share = getParkingShare(item);
  const population = numericOrNull(getPrimary("O_celk"));
  const demand = share !== null && population !== null ? Math.ceil(population * share / 100) : null;
  const section = createElement("section", "parking-share-assumption");
  const field = createElement("label", "capacity-field");
  field.append(createElement("span", "capacity-field-label", "Scénář poptávky: podíl obyvatel (%)"));
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.max = "100";
  input.step = "0.1";
  input.value = share ?? "";
  input.dataset.recordId = item.key;
  input.dataset.capacityField = "parkingShare";
  field.append(input);
  section.append(field);
  section.append(createElement("p", "parking-demand",
    demand === null ? "Zadejte přírůstek obyvatel pro výpočet scénáře poptávky."
      : `O_celk × ${formatNumber(share)} % = ${formatNumber(population)} × ${formatNumber(share / 100, 4)} = ${formatNumber(demand)} ${text(item.unit, "jednotek")} (zaokrouhlení nahoru).`));
  section.append(createElement("p", "capacity-caution",
    `${config.basis}. Scénář mění jen odhad poptávky (${formatNumber(demand)} ${text(item.unit, "jednotek")}); nepřepočítává návrhový počet, protože převod poptávky na stání v modelu závisí i na dalších neověřených předpokladech. Potřebujete-li změnit počet nebo cenu, použijte níže uživatelský návrh s poznámkou.`));
  return section;
}

function adjustedResult(item, kind = "") {
  const override = state.overrides[item.key];
  const hasUserOverride = Number.isFinite(override?.quantity) && Boolean(override.note?.trim());
  if (item.manualCoverageCheck && !hasUserOverride) {
    return {
      required: null,
      remaining: null,
      remainingCapacity: null,
      capacityDemand: null,
      capacityUnit: getCapacityUnit(item),
      usesDemandBasedCapacity: false,
      cost: null,
      verifiedCapacity: 0,
      isCapacityVerified: false,
      manualCoverageCheck: true,
      hasOverride: false
    };
  }
  const required = getPlannedQuantity(item, kind);
  if (required === null || required < 0) {
    return { required, remaining: null, remainingCapacity: null, cost: null, verifiedCapacity: 0 };
  }
  const capacity = getCapacityEdit(item);
  const verified = numericOrNull(capacity.quantity);
  const capacityVerified = verified !== null && verified >= 0;
  const isManualStandardTarget = getAppliedStandardTarget(item) !== null
    && !state.overrides[item.key]?.note?.trim();
  const capacityDemand = usesDemandBasedCapacity(item) && !isManualStandardTarget ? item.demand : required;
  const availableUse = capacityVerified && capacityDemand !== null
    ? getAllowedCapacityUse(item, verified, capacityDemand) : 0;
  const standardBaseFloor = isManualStandardTarget ? getStandardComparison(item).base : null;
  const standardUplift = standardBaseFloor === null ? null : getStandardComparison(item).uplift;
  const applied = standardUplift === null ? availableUse : Math.min(availableUse, standardUplift);
  const remainingCapacity = capacityDemand === null ? null
    : Math.max(standardBaseFloor ?? 0, capacityDemand - applied);
  const remaining = usesDemandBasedCapacity(item) ? required
    : capacityDemand === null ? null : remainingCapacity;
  const costBasis = (item.quantityBasis || usesDemandBasedCapacity(item)) ? item.demand : item.modelProposed;
  const baseUnitCost = item.cost !== null && item.cost > 0 && costBasis > 0
    ? item.cost / costBasis : null;
  return {
    required,
    remaining,
    remainingCapacity,
    capacityDemand,
    capacityUnit: getCapacityUnit(item),
    usesDemandBasedCapacity: usesDemandBasedCapacity(item),
    cost: baseUnitCost !== null ? baseUnitCost * (usesDemandBasedCapacity(item) ? remainingCapacity : remaining) : null,
    verifiedCapacity: applied,
    availableCapacity: capacityVerified ? verified : 0,
    reserveLimit: getReserveLimit(),
    isCapacityVerified: capacityVerified,
    hasOverride: Boolean(state.overrides[item.key]?.note?.trim())
  };
}

function renderRequirementList(parent, kind, records) {
  parent.replaceChildren();
  const items = records.filter((item) => {
    const hasLegalRule = item.legalRules.some((rule) => !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id])));
    const isInfrastructure = INFRASTRUCTURE_ITEM_IDS.has(item.id);
    const hasPrimaryPlacement = hasLegalRule || item.recommendedRules.length > 0
      || item.id === "VYB_07B" || RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id);
    if (kind === "infrastructure") return isInfrastructure;
    if (isInfrastructure) return false;
    if (kind === "legal") {
      return item.id === RESIDENT_PARKING_ID || item.id === "VYB_17B"
        || hasLegalRule && !RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
        && (SCHOOL_CAPACITY_ITEM_IDS.has(item.id) || item.recommendedRules.length === 0);
    }
    if (kind === "recommended") return isResidentParkingItem(item)
      ? item.id === RESIDENT_PARKING_LAW_ID || item.id === RESIDENT_PARKING_RESERVE_ID
      : item.recommendedRules.length > 0
      || item.id === "VYB_07B" || RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id);
    return !RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
      && item.otherRules.length > 0 && (!hasPrimaryPlacement || REVIEW_ALTERNATIVE_IDS.has(item.id));
  });
  const selectedKeys = getRealizationPlan().selectedKeys;
  const visibleItems = state.hideUnrealizedItems
    ? items.filter((item) => selectedKeys.has(item.key))
    : items;
  const badge = kind === "legal" ? elements.legalCount
    : kind === "recommended" ? elements.recommendedCount
      : kind === "infrastructure" ? elements.infrastructureCount : elements.reviewCount;
  const linkedRules = (state.unlinkedRules[kind] || []).filter((rule) => {
    if (state.hideUnrealizedItems) return false;
    if (state.selectedCategory) return false;
    const query = state.search.toLocaleLowerCase("cs").trim();
    if (!query) return true;
    return [rule[RULE_HEADERS.id], rule[RULE_HEADERS.equipment], rule[RULE_HEADERS.equipmentId],
      rule[RULE_HEADERS.area], rule[RULE_HEADERS.description], rule[RULE_HEADERS.source]]
      .some((value) => String(value ?? "").toLocaleLowerCase("cs").includes(query));
  });
  const generalRules = kind === "legal"
    ? getGeneralObligationRules(state.hideUnrealizedItems
      ? records.filter((item) => selectedKeys.has(item.key)) : null)
    : [];
  badge.textContent = linkedRules.length
    ? `Položky: ${visibleItems.length} · ID k přiřazení: ${linkedRules.length}`
    : String(visibleItems.length);
  if (!items.length && !linkedRules.length && !generalRules.length) {
    const empty = createElement("div", "empty-list", state.hideUnrealizedItems
      ? "V tomto seznamu nejsou položky vybrané k realizaci."
      : kind === "legal"
        ? "V této skupině nebylo nalezeno pravidlo výslovně označené jako právní povinnost."
        : kind === "recommended"
          ? "V této skupině nebylo nalezeno pravidlo výslovně označené jako metodické doporučení."
          : kind === "infrastructure"
            ? "V této skupině nejsou položky technické infrastruktury."
            : "V této skupině nejsou další nezařazená pravidla.");
    empty.dataset.realizationEmpty = "";
    empty.dataset.requirementKind = kind;
    parent.append(empty);
    return;
  }
  const sortItems = (left, right) => {
    const listSatisfied = (item) => kind === "recommended"
      && getPrimary("REZIM_STD") !== "STANDARD"
      && hasStandardComparison(item)
      && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
      ? getStandardComparison(item).delta === 0
      : isRequirementSatisfied(adjustedResult(item, kind));
    const leftZero = listSatisfied(left);
    const rightZero = listSatisfied(right);
    return Number(leftZero) - Number(rightZero);
  };
  const reserveItems = kind === "recommended"
    ? items.filter((item) => item.legalRules.some((rule) => !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id]))))
    : [];
  const regularItems = kind === "recommended"
    ? items.filter((item) => !reserveItems.includes(item))
    : items;
  if (reserveItems.length) {
    const group = createElement("section", "reserve-capacity-group");
    const heading = createElement("div", "reserve-capacity-heading");
    heading.append(createElement("h3", "", "Rezerva a navýšení kapacity"));
    heading.append(createElement("p", "", "Položky s právním minimem i metodickým doporučením; každá je zobrazena pouze jednou."));
    group.append(heading);
    reserveItems.sort(sortItems).forEach((item) => group.append(renderItemCard(item, kind)));
    parent.append(group);
  }
  regularItems.sort(sortItems).forEach((item) => parent.append(renderItemCard(item, kind)));
  if (generalRules.length) parent.append(renderGeneralObligations(generalRules));
  linkedRules.forEach((rule) => parent.append(renderUnlinkedRuleCard(rule, kind)));
  if (!visibleItems.length && !linkedRules.length && !generalRules.length) {
    const empty = createElement("div", "empty-list", "V tomto seznamu nejsou položky vybrané k realizaci.");
    empty.dataset.realizationEmpty = "";
    parent.append(empty);
  }
  updateRequirementListVisibility(parent, selectedKeys);
}

function updateRequirementListVisibility(list, selectedKeys) {
  const cards = [...list.querySelectorAll(".requirement-card")];
  if (!cards.length) {
    const empty = list.querySelector("[data-requirement-kind]");
    if (empty) {
      const kind = empty.dataset.requirementKind;
      empty.textContent = state.hideUnrealizedItems
        ? "V tomto seznamu nejsou položky vybrané k realizaci."
        : kind === "legal"
          ? "V této skupině nebylo nalezeno pravidlo výslovně označené jako právní povinnost."
          : kind === "recommended"
            ? "V této skupině nebylo nalezeno pravidlo výslovně označené jako metodické doporučení."
            : kind === "infrastructure"
              ? "V této skupině nejsou položky technické infrastruktury."
              : "V této skupině nejsou další nezařazená pravidla.";
    }
    return 0;
  }
  for (const card of cards) {
    card.hidden = state.hideUnrealizedItems && !selectedKeys.has(card.dataset.itemKey);
  }
  list.querySelectorAll(".reserve-capacity-group").forEach((group) => {
    group.hidden = !group.querySelector(".requirement-card:not([hidden])");
  });
  const visibleCards = cards.filter((card) => !card.hidden).length;
  const hasOtherContent = Boolean(list.querySelector(".general-obligations, .unlinked-rule-card"));
  const empty = list.querySelector("[data-realization-empty]");
  if (state.hideUnrealizedItems && !visibleCards && !hasOtherContent) {
    if (!empty) {
      const placeholder = createElement("div", "empty-list", "V tomto seznamu nejsou položky vybrané k realizaci.");
      placeholder.dataset.realizationEmpty = "";
      list.append(placeholder);
    }
  } else {
    empty?.remove();
  }
  return visibleCards;
}

function updateRequirementVisibility() {
  const selectedKeys = getRealizationPlan().selectedKeys;
  const visibleItems = state.hideUnrealizedItems
    ? matchingItems().filter((item) => selectedKeys.has(item.key))
    : null;
  const generalRules = getGeneralObligationRules(visibleItems);
  const existingGeneralRules = elements.legalList.querySelector(":scope > .general-obligations");
  if (generalRules.length) {
    const updatedGeneralRules = renderGeneralObligations(generalRules);
    if (existingGeneralRules) existingGeneralRules.replaceWith(updatedGeneralRules);
    else elements.legalList.append(updatedGeneralRules);
  } else {
    existingGeneralRules?.remove();
  }
  for (const [list, badge] of [
    [elements.legalList, elements.legalCount],
    [elements.infrastructureList, elements.infrastructureCount],
    [elements.recommendedList, elements.recommendedCount],
    [elements.reviewList, elements.reviewCount]
  ]) {
    const visibleCards = updateRequirementListVisibility(list, selectedKeys);
    const linkedRules = list.querySelectorAll(".unlinked-rule-card").length;
    badge.textContent = linkedRules
      ? `Položky: ${visibleCards} · ID k přiřazení: ${linkedRules}`
      : String(visibleCards);
  }
}

function getGeneralObligationRules(visibleItems = null) {
  const items = visibleItems || state.items
    .filter((item) => !state.selectedCategory || item.category === state.selectedCategory);
  const showShelterCard = items.some((item) => item.id === "VYB_17B");
  const rules = items
    .flatMap((item) => item.legalRules)
    .filter((rule) => GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id]))
      && !(showShelterCard && String(rule[RULE_HEADERS.id]) === "SAFE-LEG-001"));
  return [...new Map(rules.map((rule) => [String(rule[RULE_HEADERS.id]), rule])).values()];
}

function renderGeneralObligations(rules) {
  const section = createElement("details", "general-obligations");
  const heading = document.createElement("summary");
  heading.className = "general-obligations-heading";
  heading.append(createElement("h3", "", "Obecné závazky a podmíněné právní kontroly"));
  heading.append(createElement("span", "count-badge", rules.length));
  heading.append(createElement("span", "section-expand-mark", "＋"));
  section.append(heading);
  const content = createElement("div", "general-obligations-content");
  content.append(createElement("p", "general-obligations-note",
    "Nejde o automaticky odvozené množství zařízení. U každého bodu ověřte, zda se vztahuje na konkrétní projekt; podmíněná povinnost není univerzálním příkazem stavět novou vybavenost."));
  rules.forEach((rule) => {
    const disclosure = createElement("details", "general-obligation");
    const summary = document.createElement("summary");
    summary.append(createElement("strong", "", displayRuleName(rule)));
    summary.append(createElement("span", "section-expand-mark", "＋"));
    disclosure.append(summary);
    const obligationContent = createElement("div", "general-obligation-content");
    obligationContent.append(renderRule(rule, "legal"));
    if (!requiresCurrentVerification(rule)) obligationContent.append(renderManualReview(rule));
    disclosure.append(obligationContent);
    content.append(disclosure);
  });
  section.append(content);
  return section;
}

function renderUnlinkedRuleCard(rule, kind) {
  const details = createElement("details", "unlinked-rule-card");
  const summary = document.createElement("summary");
  const displayName = displayRuleName(rule);
  summary.append(createElement("strong", "", displayName));
  summary.append(createElement("span", "unlinked-badge", "Nepřiřazené ID"));
  details.append(summary);
  const body = createElement("div", "requirement-body");
  body.append(createElement("p", "unlinked-note",
    `${text(rule.linkIssue)} Jde o problém jednoznačného kódového přiřazení, ne o tvrzení, že pravidlo věcně nesouvisí s vybaveností. Dokud nebude mapování potvrzeno, nepoužiji je k automatickému odečtu kapacity ani ceny.`));
  const sourceEquipmentName = cleanDisplayName(text(rule[RULE_HEADERS.equipment], "Pravidlo bez názvu vybavenosti"));
  if (displayName !== sourceEquipmentName) {
    body.append(createElement("p", "original-name-note",
      `Původní název v sešitu: ${cleanDisplayName(text(rule[RULE_HEADERS.equipment]))}`));
  }
  body.append(renderRule(rule, kind));
  details.append(body);
  return details;
}

function displayRuleName(rule) {
  const id = String(rule[RULE_HEADERS.id] ?? "").trim();
  const names = {
    "ZDR-LEG-001": "Časová dostupnost základní zdravotní péče",
    "ZDR-SPEC-001": "Časová dostupnost specializované zdravotní péče",
    "ZDR-MET-001": "Dostupnost ambulantní zdravotní péče – metodický standard",
    "ZEL-MET-001": "Předškolní hřiště – docházková dostupnost",
    "ZEL-MET-002": "Hřiště pro mladší děti – docházková dostupnost",
    "ZEL-MET-003": "Hřiště pro mládež a dospělé – dostupnost",
    "ZEL-RTS-001": "Krytá sportovní hala – cenový podklad RTS",
    "ZDR-RTS-001": "Zdravotnické budovy – orientační cenový podklad",
    "CULT-COMM-001": "Maloobchod – lokální plánovací podmínky",
    "CULT-LAW-001": "Občanská vybavenost – právní rámec (bez obecné povinnosti stavět)",
    "SOC-LEG-001": "Sociální služba – právní rámec (oddělit službu od stavby)",
    "SAFE-MET-001": "Pokrytí jednotkou požární ochrany – ruční ověření",
    "SAFE-LEG-001": "Úkryt civilní ochrany – podmíněná kontrola",
    "ZEL-LEG-002": "Šířky přístupných veřejných prostranství a ulic",
    "TECH-LEGAL-002": "Místo pro soustřeďování komunálního odpadu u bydlení"
  };
  return names[id] || cleanDisplayName(text(rule[RULE_HEADERS.equipment], "Pravidlo bez názvu vybavenosti"));
}

function isResidentParkingItem(item) {
  return RESIDENT_PARKING_IDS.has(item.id);
}

function createResidentParkingVariants(items) {
  const source = items.find((item) => item.id === RESIDENT_PARKING_ID);
  if (!source) return [];
  return [RESIDENT_PARKING_LAW_ID, RESIDENT_PARKING_RESERVE_ID].map((id) => ({
    ...source,
    id,
    key: `${source.key}::${id}`,
    parkingSourceKey: source.key,
    name: id === RESIDENT_PARKING_LAW_ID
      ? "Rezidentské parkování (splnění zákona č. 283/2021 Sb.)"
      : "Rezidentské parkování Rezerva ( % nad rámec zákona č. 283/2021 Sb.)",
    legalRules: [],
    recommendedRules: [],
    otherRules: [],
    quantityBasis: false,
    proposed: 0,
    modelProposed: 0,
    demand: 0
  }));
}

function reviewKey(rule) {
  return [rule[RULE_HEADERS.id], rule[RULE_HEADERS.equipmentId], rule[RULE_HEADERS.area]]
    .map((value) => String(value ?? "").trim()).join("::");
}

function renderAlternativeContext(item) {
  const alternative = REVIEW_ALTERNATIVE_LABELS[item.id];
  if (!alternative) return null;
  const note = createElement("section", "review-placement-note");
  note.append(createElement("h4", "", "Proč je tato alternativa samostatná"));
  note.append(createElement("p", "", alternative.context));
  return note;
}

function renderTransferredReviewWarnings(item, displayedRules = []) {
  const displayedRuleIds = new Set(displayedRules
    .map((rule) => String(rule[RULE_HEADERS.id] ?? "").trim())
    .filter(Boolean));
  const seenRules = new Set();
  const rules = item.otherRules.filter((rule) => {
    const id = String(rule[RULE_HEADERS.id] ?? "").trim();
    const key = id || [
      rule[RULE_HEADERS.type],
      rule[RULE_HEADERS.description],
      rule[RULE_HEADERS.condition],
      rule[RULE_HEADERS.limitation]
    ].map((value) => String(value ?? "").trim().toLocaleLowerCase("cs")).join("::");
    if (!key || seenRules.has(key) || id && displayedRuleIds.has(id)) return false;
    seenRules.add(key);
    return true;
  });
  if (!rules.length) return null;
  const note = createElement("details", "item-section supplemental-rules-section");
  addDisclosureSummary(note, "Doplňující pravidla a upozornění", `${rules.length} pravidel`);
  const rulesContent = createElement("div", "rules-content supplemental-rules-content");
  for (const rule of rules) {
    const type = String(rule[RULE_HEADERS.type] ?? "").trim();
    if (!type) continue;
    const disclosure = createElement("details", "rule-disclosure");
    const summary = document.createElement("summary");
    summary.append(createElement("strong", "", type));
    if (rule[RULE_HEADERS.id]) summary.append(createElement("code", "rule-id", rule[RULE_HEADERS.id]));
    summary.append(createElement("span", "section-expand-mark", "＋"));
    disclosure.append(summary);
    const ruleContent = createElement("div", "supplemental-rule-content");
    const safeField = (field) => safeWorkbookDisplayText(rule[field]);
    const description = safeField(RULE_HEADERS.description);
    const condition = safeField(RULE_HEADERS.condition);
    const verification = safeField(RULE_HEADERS.verification);
    const limitation = safeField(RULE_HEADERS.limitation);
    if (description) addField(ruleContent, "Co pravidlo stanoví", description);
    if (condition) addField(ruleContent, "Podmínka / rozsah", condition);
    if (verification) addField(ruleContent, "Stav ověření", verification);
    if (limitation) addField(ruleContent, "Omezení / nutné rozhodnutí", limitation);
    const citation = [rule[RULE_HEADERS.source], rule[RULE_HEADERS.location]]
      .map(readableWorkbookCitation)
      .filter((value) => value && !isWorkbookImplementationText(value))
      .join(" · ");
    if (citation) addField(ruleContent, "Zdroj", citation, "citation-field");
    if (!ruleContent.childElementCount) continue;
    disclosure.append(ruleContent);
    rulesContent.append(disclosure);
  }
  if (!rulesContent.childElementCount) return null;
  note.append(rulesContent);
  return note;
}

function isWorkbookImplementationText(value) {
  return /(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\b[A-Z]{1,3}\$?\d+\b|\b(?:ROUND(?:UP|DOWN)?|IF|SUM|MAX|MIN|CEILING|FLOOR|SQRT|VLOOKUP|XLOOKUP)\s*\(|\b[A-Za-z]{1,3}_[A-Za-z0-9_]+\b|^\s*=|Model_Vyb|Cluster_[A-Za-z0-9_]+/i
    .test(String(value ?? ""));
}

function safeWorkbookDisplayText(value) {
  const cleanValue = String(value ?? "").trim();
  return cleanValue && !isWorkbookImplementationText(cleanValue) ? cleanValue : "";
}

function getReviewAlternativeName(item) {
  return REVIEW_ALTERNATIVE_LABELS[item.id]?.name || displayItemName(item);
}

function reviewGuidance(rule) {
  const id = String(rule[RULE_HEADERS.id] ?? "").trim();
  const guidance = {
    "ZDR-LEG-001": {
      summary: "Jde o kontrolu dostupnosti hrazené služby, ne o automatický požadavek postavit ordinaci. Rozhodující skupina a limit musí odpovídat konkrétní odbornosti a typu péče.",
      checks: [
        "Určete odbornost a skupinu podle citované přílohy; ověřte, zda se jedná o ambulantní hrazenou péči.",
        "Ověřte skutečnou dobu dojezdu ke službě a použitý způsob dopravy; limity v záznamu jsou 35 / 45 / 60 / 90 / 120 minut podle skupiny.",
        "Zjistěte, zda služba přijímá pacienty a zda je dostupná pro relevantní pojištěnce. Samotný čas dojezdu neprokazuje volnou kapacitu.",
        "Zapište, zda je požadavek pro záměr pouze kontextem zdravotní dostupnosti, nebo zda byl doložen zvláštní závazný požadavek."
      ]
    },
    "ZDR-SPEC-001": {
      summary: "Nařízení vlády č. 307/2012 Sb. stanoví nejzazší dojezdové časy pro hrazené služby, nikoli povinnost postavit místní ordinaci. V ambulantní péči jsou chirurgie, vnitřní lékařství, ORL a oftalmologie ve skupině 2 (45 minut); kardiologie a dermatovenerologie ve skupině 3 (60 minut). Neurologie dospělých do těchto skupin automaticky nepatří.",
      checks: [
        "Sepište konkrétní ambulantní odbornost a ověřte její skupinu/limit v citovaném zdroji.",
        "Ověřte reálnou dobu dojezdu, poskytování hrazené péče a zda pracoviště přijímá pacienty.",
        "Neurologii dospělých posuďte podle konkrétního typu služby; příloha samostatně uvádí dětskou neurologii v jiné skupině a limitech.",
        "Uveďte, zda ověřujete dostupnost služby, nebo kapacitu; jde o různé údaje."
      ]
    },
    "ZEL-MET-001": {
      summary: "V tomto záznamu se nemá používat odhadný koeficient místo sítě: metodický práh je skutečná pěší trasa. Nejde o zákonný příkaz postavit hřiště.",
      checks: [
        "Určete vstupní body bydlení a veřejně přístupný vstup hřiště.",
        "Změřte skutečnou průchozí trasu v pěší síti; zaznamenejte délku a GIS/mapový podklad.",
        "Zkontrolujte bariéry, bezpečné přechody, uzavírky a veřejný přístup. Vzdušná vzdálenost není náhradou trasy.",
        "Porovnejte trasu s doporučením 200 m uvedeným ve zdroji; standard nestanoví kapacitu hřiště."
      ]
    },
    "CULT-COMM-001": {
      summary: "Maloobchod je komerční vybavenost. V katalogu není doložen obecný český ani TAČR/MMR závazný poměr m² prodejní plochy na obyvatele. Územní plán nebo regulační podmínky mohou mít místní, konkrétní pravidla; ta se musí ověřit v platné dokumentaci pro daný pozemek.",
      checks: [
        "Vyhledejte platný Územní plán města Brna a případné regulační podmínky pro konkrétní lokalitu; zapište přesné označení plochy a ustanovení.",
        "Oddělte závazný limit nebo podmínku využití plochy od obecné strategie, metodiky, komerční studie či tržního odhadu.",
        "Nezadávejte m² prodejní plochy na obyvatele jako normu, dokud není nalezen konkrétní platný zdroj a jeho působnost.",
        "Zaznamenejte stávající obchody, spádové území, otevírací dobu a případné sdílené kapacity; samostatnou prodejnu zvažujte jako projektový scénář."
      ]
    },
    "SAFE-MET-001": {
      summary: "Tato položka ověřuje pokrytí jednotkou požární ochrany, nikoli automatickou potřebu nové stanice nebo zbrojnice. Pěší dostupnost není měřítkem zásahového dojezdu.",
      checks: [
        "Zjistěte, které jednotky (HZS/JPO I a JSDH/JPO II–V) pokrývají lokalitu podle aktuálního požárního poplachového plánu Jihomoravského kraje.",
        "Ověřte kategorii, územní působnost, výjezduschopnost, techniku a skutečný dojezd po relevantních trasách; uveďte datum a mapový/plánovací podklad.",
        "Nepoužívejte docházku 1 000 m ani počet obyvatel jako důkaz povinnosti zřídit novou stanici.",
        "Případnou potřebu nové stanice nebo zbrojnice doložte samostatným stanoviskem HZS/obce a příslušnou koncepcí; HZS a JSDH nejsou zaměnitelné."
      ]
    },
    "ZEL-LEG-002": {
      summary: "Jde o kontrolu parametrů veřejných prostranství a přístupových ulic, pokud je projekt navrhuje; samo pravidlo neurčuje počet zařízení ani automatickou novou investici do vybavenosti.",
      checks: [
        "Určete, které nové nebo měněné ulice a veřejná prostranství jsou součástí záměru.",
        "Ověřte předepsanou šířku podle platného právního předpisu, charakteru ulice a konkrétního návrhu; zaznamenejte měření a výkres.",
        "Rozlište veřejný uliční prostor od započitatelné plochy veřejného prostranství; komunikace se do plošného minima automaticky nezapočítává.",
        "Uveďte zdroj, verzi dokumentace a případné odůvodněné výjimky."
      ]
    },
    "SAFE-LEG-001": {
      summary: "Požadavek na úkryt civilní ochrany je podmíněný. Záznam neznamená povinnost vybudovat stálý úkryt v každém projektu.",
      checks: [
        "Ověřte, zda se pro lokalitu a záměr uplatní požadavky ochrany obyvatelstva podle aktuálních právních předpisů a dokumentace obce/kraje.",
        "Zkontrolujte, zda územní nebo krizová dokumentace stanoví konkrétní úkryt, kapacitu, dosah či jiné opatření.",
        "Zapište odpovědný orgán, stanovisko a přesný zdroj; pokud se povinnost neuplatní, uveďte důvod.",
        "Nevyvozujte potřebu stálého úkrytu pouze z počtu nových obyvatel ani z nezávazného zahraničního standardu."
      ]
    },
    "TECH-LEGAL-002": {
      summary: "Jde o požadavek na místo/prostor pro soustřeďování komunálního odpadu u dotčené stavby pro bydlení; konkrétní uspořádání a případné společné stanoviště se ověřuje pro projekt. Nejde o výpočet nové sběrny nebo sběrného dvora.",
      checks: [
        "V projektové dokumentaci určete místo pro nádoby a přístup k němu pro obyvatele i svozovou techniku.",
        "Ověřte aktuální požadavky vyhlášky č. 146/2024 Sb. a zákona č. 541/2020 Sb., včetně podmínek případného společného stanoviště.",
        "Zaznamenejte kapacitu/nádoby, vlastníka nebo správce, svozový režim, bezpečný přístup a konkrétní výkres či vyjádření.",
        "Rozlište povinné místo u stavby od doporučené docházkové vzdálenosti ke stanovišti tříděného odpadu nebo od sběrného dvora."
      ]
    }
  };
  return guidance[id] || {
    summary: "Pravidlo není jednoznačně přiřazeno k položce modelu. Ověření uložte samostatně; aplikace z něj bez potvrzení nevypočítává počet ani cenu.",
    checks: [
      "Ověřte, zda se pravidlo a citovaný zdroj vztahují na tento záměr a na daný typ území.",
      "Zjistěte, které vstupy, územní rozsah a jednotky jsou pro kontrolu potřeba; nevyplněné parametry označte.",
      "Ověřte skutečný stav v území a případnou využitelnou kapacitu, pokud ji pravidlo řeší.",
      "Rozlište právní povinnost, metodické doporučení a informativní/cenový podklad podle uvedeného zdroje."
    ]
  };
}

function renderManualReview(rule) {
  const id = reviewKey(rule);
  const saved = state.reviewChecks[id] || { status: "Nezahájeno", source: "", finding: "", note: "" };
  const guidance = reviewGuidance(rule);
  const section = createElement("section", "manual-review");
  section.dataset.reviewId = id;
  section.append(createElement("h4", "", "Kontrola uživatelem"));
  section.append(createElement("p", "manual-review-summary", guidance.summary));
  const list = createElement("ul", "manual-review-checklist");
  guidance.checks.forEach((check) => list.append(createElement("li", "", check)));
  section.append(list);

  const fields = createElement("div", "capacity-controls");
  const statusLabel = createElement("label", "capacity-field");
  statusLabel.append(createElement("span", "capacity-field-label", "Stav kontroly"));
  const status = document.createElement("select");
  status.dataset.reviewId = id;
  status.dataset.reviewField = "status";
  ["Nezahájeno", "Ověřeno – vyhovuje", "Ověřeno – nedostupné / deficit", "Neplatí pro záměr"].forEach((value) => {
    status.append(new Option(value, value));
  });
  status.value = saved.status;
  statusLabel.append(status);
  fields.append(statusLabel);
  fields.append(createReviewInput("Podklad / zdroj ověření", saved.source, id, "source", "Např. GIS, registr poskytovatelů, vyjádření obce"));
  fields.append(createReviewInput("Zjištěný stav / výsledek", saved.finding, id, "finding", "Doplňte naměřenou hodnotu, kapacitu nebo co chybí"));
  fields.append(createReviewInput("Poznámka / další krok", saved.note, id, "note", "Včetně rozsahu, data a nejistot"));
  section.append(fields);
  return section;
}

function createReviewInput(label, value, id, field, placeholder) {
  const wrapper = createElement("label", "capacity-field");
  wrapper.append(createElement("span", "capacity-field-label", label));
  const input = document.createElement(field === "note" ? "textarea" : "input");
  if (field === "note") input.rows = 2;
  input.value = value ?? "";
  input.placeholder = placeholder;
  input.dataset.reviewId = id;
  input.dataset.reviewField = field;
  wrapper.append(input);
  return wrapper;
}

function getHealthChecks(item) {
  const saved = state.healthChecks[item.key];
  if (Array.isArray(saved)) return saved;
  return (HEALTH_SPECIALTIES[item.id] || []).map((specialty, index) => ({
    id: `default-${index}`,
    profession: specialty.profession,
    group: specialty.group,
    limitMinutes: specialty.limitMinutes ?? "",
    provider: "",
    travelMinutes: "",
    capacity: "",
    demand: "",
    source: "",
    note: ""
  }));
}

function healthSearchUrl(item, entry) {
  const place = String(getPrimary(PROJECT_LOCATION_INPUT.code) || PROJECT_LOCATION_INPUT.defaultValue).trim();
  const population = derivedValues();
  const populationText = (value) => value === null || value === undefined
    ? "nezadán" : `${formatNumber(value)} obyvatel`;
  const populationContext = [
    `přírůstek projektu: ${populationText(getPrimary("O_celk"))}`,
    `současný počet obyvatel řešeného území: ${populationText(getPrimary("O_sidlo"))}`,
    `okolní spádující sídla: ${populationText(getPrimary("O_spad"))}`,
    `celkem po záměru a ve spádu: ${populationText(population.O_prah)}`
  ].join("; ");
  const specialtyQuery = entry?.profession
    ? `${entry.profession}${entry.group ? ` (skupina ${entry.group})` : ""}`
    : getHealthChecks(item).map((record) =>
      `${record.profession}${record.group ? ` (skupina ${record.group}` : ""}${record.group ? `${record.limitMinutes ? `, ${record.limitMinutes} min` : ""})` : ""}`)
      .filter(Boolean).join(", ");
  const query = [
    `Ověřte zdravotní dostupnost v lokalitě ${place}`,
    specialtyQuery || displayItemName(item),
    entry?.group ? `skupina ${entry.group}` : "",
    entry?.limitMinutes ? `limit ${entry.limitMinutes} minut` : "",
    populationContext,
    "najděte poskytovatele, skutečnou dobu dojezdu, zda přijímá pacienty, doloženou kapacitu a zdroje",
    "odhad potřeby nebo deficitu uveďte pouze při doložené metodice a podkladech"
  ].filter(Boolean).join("; ");
  return `https://www.google.com/search?q=${encodeURIComponent(query)}`;
}

function healthCapacitySummary(entry) {
  const nonNegative = (value) => {
    const number = numericOrNull(value);
    return number !== null && number >= 0 ? number : null;
  };
  const capacity = nonNegative(entry.capacity);
  const demand = nonNegative(entry.demand);
  const travel = nonNegative(entry.travelMinutes);
  const limit = nonNegative(entry.limitMinutes);
  const hasSource = Boolean(entry.source?.trim());
  const summary = [];
  if (travel !== null && limit !== null) {
    summary.push(travel <= limit
      ? `Zadaný dojezd ${formatNumber(travel)} min nepřekračuje vložený limit ${formatNumber(limit)} min.`
      : `Zadaný dojezd ${formatNumber(travel)} min překračuje vložený limit ${formatNumber(limit)} min.`);
    if (!hasSource) summary.push("Doložte podklad do pole zdroje.");
  }
  if (capacity !== null && demand !== null && hasSource) {
    const difference = capacity - demand;
    summary.push(difference >= 0
      ? `Zadaná kapacita převyšuje zadanou potřebu o ${formatNumber(difference)} osob.`
      : `Zadaný deficit je ${formatNumber(Math.abs(difference))} osob.`);
  } else {
    summary.push(capacity !== null && demand !== null
      ? "Kapacitní rezervu/deficit nelze vypočítat bez doloženého zdroje."
      : "Kapacitní rezervu/deficit nelze vypočítat bez doložené potřeby i kapacity.");
  }
  return summary.join(" ");
}

function createHealthField(item, entry, field, label, options = {}) {
  const wrapper = createElement("label", `capacity-field health-field${options.wide ? " health-field-wide" : ""}`);
  wrapper.append(createElement("span", "capacity-field-label", label));
  const input = document.createElement("input");
  input.type = options.numeric ? "number" : "text";
  if (options.numeric) {
    input.min = "0";
    input.step = "any";
    input.inputMode = "decimal";
  }
  input.value = entry[field] ?? "";
  input.placeholder = options.placeholder || "";
  input.dataset.healthItem = item.key;
  input.dataset.healthId = entry.id;
  input.dataset.healthField = field;
  wrapper.append(input);
  return wrapper;
}

function renderHealthCheckRecord(item, entry) {
  const record = createElement("article", "health-check-record");
  record.dataset.healthRecord = entry.id;
  const top = createElement("div", "health-check-top");
  top.append(createElement("strong", "", entry.profession || "Nová odbornost"));
  const link = createElement("a", "source-link health-search-link", "Hledat obor ↗");
  link.href = healthSearchUrl(item, entry);
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  top.append(link);
  record.append(top);

  const fields = createElement("div", "health-fields");
  fields.append(createHealthField(item, entry, "profession", "Odbornost / služba", { placeholder: "Např. neurologie" }));
  fields.append(createHealthField(item, entry, "group", "Skupina / druh péče", { placeholder: "Ověřit ve zdroji" }));
  fields.append(createHealthField(item, entry, "limitMinutes", "Limit dojezdu (min)", { numeric: true, placeholder: "Neznámý = nechat prázdné" }));
  fields.append(createHealthField(item, entry, "provider", "Poskytovatel / pracoviště", { placeholder: "Název zařízení" }));
  fields.append(createHealthField(item, entry, "travelMinutes", "Ověřený dojezd (min)", { numeric: true }));
  fields.append(createHealthField(item, entry, "capacity", "Ověřená kapacita (osob)", { numeric: true }));
  fields.append(createHealthField(item, entry, "demand", "Potřeba ze zdroje (osob)", { numeric: true }));
  fields.append(createHealthField(item, entry, "source", "Zdroj / datum ověření", { wide: true, placeholder: "Registr, pojišťovna, GIS, vyjádření poskytovatele" }));
  fields.append(createHealthField(item, entry, "note", "Poznámka", { wide: true, placeholder: "Např. přijímá nové pacienty, rozsah služby" }));
  record.append(fields);
  record.append(createElement("p", "health-check-result", healthCapacitySummary(entry)));
  return record;
}

function renderHealthCheckEditor(item) {
  if (!HEALTH_SPECIALTIES[item.id] || !isHealthcareItem(item)) return null;
  const disclosure = createElement("details", "item-section health-section");
  addDisclosureSummary(disclosure, "Dostupnost zdravotních služeb – kontrola po odbornostech");
  const section = createElement("section", "health-check-editor");
  section.dataset.healthEditor = item.key;
  const heading = createElement("div", "health-check-heading");
  const allLink = createElement("a", "text-button health-search-all", "Vyhledat všechny údaje k položce ↗");
  allLink.href = healthSearchUrl(item);
  allLink.target = "_blank";
  allLink.rel = "noopener noreferrer";
  heading.append(allLink);
  section.append(heading);
  section.append(createElement("p", "health-check-note",
    "Záznam neznamená automatickou povinnost vybudovat ordinaci. Ověřte poskytovatele a dostupnost v registrech, u pojišťoven a přímo u pracoviště. Kapacitu ani potřebu aplikace neodhaduje bez doložených čísel; limity a skupiny ověřte v citovaném zdroji pro konkrétní službu."));
  const records = createElement("div", "health-check-records");
  getHealthChecks(item).forEach((entry) => records.append(renderHealthCheckRecord(item, entry)));
  section.append(records);
  const addButton = createElement("button", "text-button health-add-button", "+ Přidat odbornost / pracoviště");
  addButton.type = "button";
  addButton.dataset.addHealthItem = item.key;
  section.append(addButton);
  disclosure.append(section);
  return disclosure;
}

function displayItemName(item) {
  if (TECHNICAL_CATALOG_NAMES[item.id]) return TECHNICAL_CATALOG_NAMES[item.id];
  if (item.id === "VYB_16A") return "Hasičská stanice HZS – pokrytí a případná potřeba";
  if (item.id === "VYB_16B") return "Jednotka požární ochrany – ruční ověření pokrytí";
  if (item.id === RESIDENT_PARKING_ID) return "Rezidentské parkování (splnění vyhlášky č. 146/2024 Sb.)";
  if (item.id === RESIDENT_PARKING_LAW_ID) return "Rezidentské parkování (splnění zákona č. 283/2021 Sb.)";
  if (item.id === RESIDENT_PARKING_RESERVE_ID) return "Rezidentské parkování Rezerva ( % nad rámec zákona č. 283/2021 Sb.)";
  const name = cleanDisplayName(ITEM_DISPLAY_NAMES[item.id] || item.name);
  const assumption = PARKING_SHARE_ASSUMPTIONS[item.id];
  return assumption
    ? `${name} (odhad pro ${formatNumber(getParkingShare(item))} % obyvatel)`
    : name;
}

function displayCatalogItemName(item) {
  return TECHNICAL_CATALOG_NAMES[item.id] || displayItemName(item);
}

function cleanDisplayName(value) {
  return String(value ?? "")
    .replace(/\s*\((?:VYB[_-]?[A-Z0-9]+|\d+(?:\.\d+)+)\)\s*/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function addDisclosureSummary(disclosure, title, detailText = "") {
  const summary = document.createElement("summary");
  const label = createElement("strong", "", title);
  summary.append(label);
  if (detailText) summary.append(createElement("small", "section-summary-detail", detailText));
  summary.append(createElement("span", "section-expand-mark", "＋"));
  disclosure.append(summary);
}

function renderItemMetrics(item, kind, isDelta = false) {
  const metrics = createElement("section", "item-results");
  const calc = adjustedResult(item, kind);
  const housingEstimate = getHousingContributionEstimate(item);
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  if (projectEstimate?.type === "targeted") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const projectMetrics = renderProjectEstimateMetrics(item, kind);
    if (projectMetrics) metrics.append(projectMetrics);
    metrics.append(createElement("p", "item-results-note",
      "Návrhová poptávka, kapacity, výměry a orientační investice jsou dopočtené ze zobrazených scénářových předpokladů. Ověřená okolní nabídka a skutečná dostupnost mohou potřebu snížit; scénář sám neurčuje právní povinnost ani nenahrazuje projekt a rozpočet."));
    return metrics;
  }
  if (projectEstimate?.type === "public-space") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const grid = createElement("div", "item-results-grid");
    const addMetric = (label, value, unit) => {
      const metric = createElement("div", "item-result-metric");
      metric.append(createElement("span", "", label));
      metric.append(createElement("strong", "", formatValue(value, unit)));
      grid.append(metric);
    };
    addMetric("Kvalifikovaná zastavitelná plocha", projectEstimate.qualifyingAreaHa, "ha");
    addMetric("Minimální výměra veřejného prostranství", projectEstimate.minimumArea, "m²");
    metrics.append(grid, createElement("p", "item-results-note",
      "Započítá se pouze související zastavitelná plocha bydlení, rekreace, občanského vybavení nebo smíšená obytná; pozemní komunikace se do minima nezahrnují. Zkontrolujte skladbu ploch v územním plánu."));
    return metrics;
  }
  if (projectEstimate?.type === "item" && item.id === "VYB_10B_1") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const projectMetrics = renderProjectEstimateMetrics(item, kind);
    if (projectMetrics) metrics.append(projectMetrics);
    const rateNote = projectEstimate.rateSource === "RTS"
      ? `${RTS_801_SOURCE} Cena se vybírá podle konstrukční charakteristiky JKSO 801.4.`
      : projectEstimate.rateSource === "missing"
        ? "RTS pro zvolenou konstrukční charakteristiku sazbu neuvádí; předpoklad lze doplnit ručně."
        : "Uživatelsky zadaná sazba není ověřenou cenou RTS.";
    metrics.append(createElement("p", "item-results-note",
      `Při více než 1 000 obyvatelích metodika doporučuje prověřit existenci a dostupnou kapacitu sálu; nejde o povinnost postavit nový objekt. Výchozí scénář předpokládá jeden sál a jeho velikost je upravitelná. ${rateNote}`));
    return metrics;
  }
  if (projectEstimate?.type === "polyclinic") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const projectMetrics = renderProjectEstimateMetrics(item, kind);
    if (projectMetrics) metrics.append(projectMetrics);
    metrics.append(createElement("p", "item-results-note",
      `Jde o poměrný scénář podle ${formatNumber(projectEstimate.population)} obyvatel a pracovního předpokladu ${formatNumber(projectEstimate.peoplePerCentre)} obyvatel na centrum. Excelový modelový počet zařízení zůstává nulový, protože tato hranice není dosažena; neznamená to nulovou projektovou plochu.`));
    return metrics;
  }
  if (!isResidentParkingItem(item)) {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
  }
  const grid = createElement("div", "item-results-grid");
  const scale = isDelta ? getStandardMetricScale(item, "recommended", calc)
    : getStandardMetricScale(item, kind, calc);
  const addMetric = (label, value, unit = "") => {
    const metric = createElement("div", "item-result-metric");
    metric.append(createElement("span", "", label));
    metric.append(createElement("strong", "", value === null || value === undefined
      ? "Neuvedeno v sešitu" : formatValue(value, unit)));
    grid.append(metric);
  };
  if (housingEstimate) {
    addMetric("Byty v bytových domech v záměru", housingEstimate.totalApartments, "bytů");
    addMetric("Byty odkupované městem za nákladovou cenu", housingEstimate.municipalApartments, "bytů");
    addMetric("Modelová hodnota celého bytového portfolia", housingEstimate.portfolioValue, "Kč");
    addMetric("Modelová hodnota městského podílu", housingEstimate.municipalValue, "Kč");
    addMetric("Odhad ušlé marže", housingEstimate.lostMargin, "Kč");
    metrics.append(grid);
    metrics.append(createElement("p", "item-results-note",
      "Výpočet zahrnuje pouze byty v bytových domech dle typové skladby A–D; rodinné domy se do městského odkupu nezapočítávají. Hodnota portfolia je orientační přepočet obestavěného prostoru sazbou 12 400 Kč/m³, nikoli odhad tržní prodejní ceny. Podrobný postup a zdroje jsou uvedeny ve výpočtu níže."));
    return metrics;
  }
  if (item.manualCoverageCheck) {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    addMetric("Výstup modelu", "Ruční ověření pokrytí", "");
    metrics.append(grid);
    metrics.append(createElement("p", "item-results-note",
      "Návrhová hodnota „1“ v pracovním modelu není počet požárních stanic ani příkaz k jejich výstavbě. Ověřte existující jednotku, její kategorii a zásahové pokrytí v požárním poplachovém plánu; případnou novou stanici posuďte samostatně podle příslušného orgánu."));
    return metrics;
  }
  if (isResidentParkingItem(item)) {
    const legal = getResidentParkingLegalEstimate(item);
    const recommended = getResidentParkingRecommendedEstimate(item);
    if (item.id === RESIDENT_PARKING_ID) {
      metrics.append(createElement("h4", "", RESULTS_HEADING));
      addMetric("Zákonné minimum po váženém součtu", legal.requiredStalls ?? "Chybí vstup",
        legal.requiredStalls === null ? "" : "stání");
      addMetric(legal.manualFloorArea === null
        ? "Modelově odhadnutá podlahová plocha bydlení"
        : "Ručně zadaná podlahová plocha dle právní definice",
      legal.floorArea ?? "Chybí vstup",
        legal.floorArea === null ? "" : "m²");
      addMetric("Běžné bydlení · plocha / 120 m²", legal.ordinaryStalls ?? "Chybí vstup",
        legal.ordinaryStalls === null ? "" : "stání");
      addMetric("Městský podíl dostupného nájemního bydlení · plocha / 240 m²",
        legal.affordableRentalStalls ?? "Chybí vstup",
        legal.affordableRentalStalls === null ? "" : "stání");
      addMetric("Plocha pro dostupné nájemní bydlení", legal.affordableRentalArea ?? "Chybí vstup",
        legal.affordableRentalArea === null ? "" : "m²");
      addMetric("Podíl bytů odkupovaných městem", legal.buyoutShare * 100, "%");
      addMetric("Korekce minima dle místního plánu", legal.minimumFactor, "koeficient");
      metrics.append(grid);
      metrics.append(createElement("p", "item-results-note",
        "Výpočet dle vyhlášky č. 146/2024 Sb. používá sazbu 1 stání / 120 m². Ručně zadaná podlahová plocha dle právní definice má přednost; jinak je plocha odhadnuta z HPP nebo počtu jednotek. HPP i průměrné plochy jsou pouze aproximační podklady."));
    } else if (item.id === RESIDENT_PARKING_LAW_ID) {
      metrics.append(createElement("h4", "", RESULTS_HEADING));
      addMetric("Modelovaný počet stání pro smíšenou zástavbu", recommended?.modeledStalls ?? "Chybí vstup",
        recommended === null ? "" : "stání");
      addMetric("Stání splňující vyhlášku č. 146/2024 Sb.", legal.requiredStalls ?? "Chybí vstup",
        legal.requiredStalls === null ? "" : "stání");
      addMetric("Doplňková stání dle zákona č. 283/2021 Sb.",
        recommended?.additionalStalls ?? "Chybí vstup",
        recommended === null ? "" : "stání");
      metrics.append(grid);
      metrics.append(createElement("p", "item-results-note",
        recommended === null
          ? "Doplňte počet bytů, HPP nebo ruční zákonnou plochu."
          : `Výsledek je pouze kladný rozdíl: max(0, ${formatNumber(recommended.modeledStalls)} modelových − ${formatNumber(legal.requiredStalls)} stání dle vyhlášky) = ${formatNumber(recommended.additionalStalls)} stání. Při realizaci obou položek se první počet započte jednou a tato položka přidá pouze rozdíl, nikdy záporný počet.`));
    } else {
      metrics.append(createElement("h4", "", RESULTS_HEADING));
      addMetric("Základ po započtení zákonného minima a doplňkových stání",
        recommended?.combinedBaseStalls ?? "Chybí vstup",
        recommended === null ? "" : "stání");
      addMetric("Procento rezervy", recommended?.reservePercent === undefined
        ? "Chybí vstup" : recommended.reservePercent * 100, "%");
      addMetric("Rezerva k samostatné realizaci", recommended?.reserveStalls ?? "Chybí vstup",
        recommended === null ? "" : "stání");
      metrics.append(grid);
      metrics.append(createElement("p", "item-results-note",
        recommended === null
          ? "Doplňte počet bytů, HPP nebo ruční zákonnou plochu."
          : `Rezerva je samostatná položka: zaokrouhlení nahoru (${formatNumber(recommended.combinedBaseStalls)} × ${formatNumber(recommended.reservePercent * 100)} %) = ${formatNumber(recommended.reserveStalls)} stání. Základ tvoří zákonné minimum a pouze nezáporný rozdíl z druhé parkovací varianty.`));
    }
    return metrics;
  }
  if (item.demand !== null) {
    addMetric(DEMAND_LABELS[item.id] ? "Počet uživatelů / míst" : getDemandLabel(item),
      getStandardMetricValue(item.demand, scale, getCapacityUnit(item)), getDemandUnit(item));
  }
  if (isHealthcareStaffingItem(item)) {
    const staffingBasis = getHealthcareStaffingBasis(item);
    metrics.append(createElement("p", "item-results-note",
      `Jde o modelový přírůstek personální potřeby, nikoli o kapacitu budovy ani „přírůstek z původní kapacity“. Sešit pro ${staffingBasis.label} dělí přírůstek O_celk poměrem ${formatNumber(staffingBasis.value)} obyvatel na jednotku souč. zátěže; rozdílné hodnoty VPL, PLDD a stomatologie proto vycházejí z různých jmenovatelů. Výchozí hodnoty a jejich indikativní povaha jsou popsány v pracovním zdravotnickém clusteru a číselníku; nejde o právní kvótu. U PLDD se používá celá populace, nikoli věkově vymezená dětská skupina.`));
  }
  if (["VYB_09B_1", "VYB_09B_2"].includes(item.id)) {
    metrics.append(createElement("p", "item-results-note",
      "Právní základ se vztahuje k místní dostupnosti hrazené služby a nejzazší době dojezdu; neukládá povinnost zřídit ordinaci v řešeném území. Položka je v zákonné části kvůli pravidlu ZDR-SPEC-001. Ostatní lékařské služby jsou v doporučeních, pokud na ně navazuje samostatné metodické pravidlo ZDR-MET-001."));
  }
  if (item.id === "VYB_14") {
    metrics.append(createElement("p", "item-results-note",
      "V sešitu je položka zastávek veřejné dopravy (VYB_14) a doporučení MOB-MET-001. Posuzuje pěší trasu ke středu nástupiště; modelové 1–2 zastávky nejsou závazný počet ani právní kvóta."));
  }
  if (item.modelProposed !== null) {
    const usesModelCountAsReference = isDelta || hasAppliedStandardTarget(item);
    addMetric(usesModelCountAsReference
      ? "Počet zařízení / bodů v základním modelu" : "Počet zařízení / bodů N_návrh",
    usesModelCountAsReference ? item.modelProposed
      : getStandardMetricValue(item.modelProposed, scale, modelCountUnit(item)), modelCountUnit(item));
  }
  addMetric("Užitná plocha", getStandardMetricValue(item.usableArea, scale), "m²");
  addMetric("Plocha pozemku", getStandardMetricValue(item.parcelArea, scale), "m²");
  addMetric("Obestavěný prostor", getStandardMetricValue(item.volume, scale), "m³");
  if (item.pavedArea !== null) addMetric("Odhad zpevněné plochy areálu",
    getStandardMetricValue(item.pavedArea, scale), "m²");
  addMetric("Odhad investice pro zvolený výsledek", getDisplayedCost(item, calc, kind), "Kč");
  metrics.append(grid);
  const sourceMetricsNote = item.id === "VYB_20" && isHealthcareItem(item)
    ? "Základní hodnoty 0 v Excelu odpovídají tomu, že ve výchozím scénáři nebyl dosažen populační spouštěč centra. Poměrný propočet polikliniky je připojen k těmto výsledkům a jeho předpoklady jsou v navazujícím výpočtu."
    : isDelta || hasAppliedStandardTarget(item)
      ? `Koeficient f = ${formatNumber(scale, 3)} škáluje plochy, kubaturu a investici. Počet zařízení zůstává referencí základního modelu; procentní navýšení kapacity samo neurčuje počet nových budov. ${item.usableArea === null || item.volume === null ? "Chybějící údaje aplikace nedopočítává z neověřené výšky ani konstrukčního předpokladu." : "Plochy a kubatura jsou převzaté z propojených listů clusterů."}`
      : scale === 1
      ? "Plochy a kubatura jsou převzaté z propojených listů clusterů v pracovním sešitu. „Neuvedeno“ znamená, že pro položku není hodnota v sešitu; aplikace ji nedopočítává z neověřené výšky ani konstrukčního předpokladu."
      : `Zobrazené modelové hodnoty jsou přepočteny koeficientem f = ${formatNumber(scale, 3)} podle základního množství, doporučeného navýšení a kapacity započtené z okolí. Jde o proporční projektový odhad, ne o novou kalkulaci RTS.`;
  metrics.append(createElement("p", "item-results-note", sourceMetricsNote));
  const projectMetrics = renderProjectEstimateMetrics(item, kind);
  if (projectMetrics) grid.append(...projectMetrics.childNodes);
  return metrics;
}

function workbookSourceName(sheetName) {
  const sourceNames = {
    Cluster_Doprava_a_Bezpeci: "dopravní a bezpečnostní cluster pracovního sešitu",
    Cluster_Skolstvi: "školský cluster pracovního sešitu",
    Cluster_Socialni_a_Bydleni: "cluster sociální vybavenosti a bydlení pracovního sešitu",
    Cluster_Zdravotnictvi: "zdravotnický cluster pracovního sešitu",
    Konstanty_a_Ciselniky: "číselník pracovního sešitu",
    Model_Vybavenosti: "list Model vybavenosti pracovního sešitu",
    Pravidla_a_vzorce: "list pravidel pracovního sešitu",
    Vstupy_a_KPI: "list vstupů a ukazatelů pracovního sešitu"
  };
  return sourceNames[sheetName] || "pracovní sešit";
}

function readableWorkbookCitation(value) {
  if (!value) return "";
  const referencePattern = /(?:'([^']+)'|([A-Za-z0-9_]+))!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g;
  return String(value).replace(referencePattern, (reference, quotedSheet, unquotedSheet) =>
    workbookSourceName(quotedSheet || unquotedSheet));
}

function calculationSource(label, source) {
  const address = source?.sourceAddress || source?.address;
  if (!address) return null;
  const sheetName = address.split("!")[0].replace(/^'|'$/g, "");
  return `${label}: ${workbookSourceName(sheetName)}.`;
}

function renderItemCalculation(item, kind) {
  const calc = adjustedResult(item, kind);
  const section = createElement("details", "item-section calculation-section");
  addDisclosureSummary(section, "Výpočet / jak jsme k výsledku došli");
  const content = createElement("div", "calculation-content");

  if (getProjectPlanningEstimate(item, kind)?.type === "targeted") {
    content.append(renderProjectEstimateCalculation(item, kind));
    if (!(HEALTH_SPECIALTIES[item.id] && isHealthcareItem(item))) {
      const verification = renderTargetedVerificationSection(item, kind);
      if (verification) content.append(verification);
    }
    section.append(content);
    return section;
  }

  if (item.manualCoverageCheck) {
    const calculation = renderProjectEstimateCalculation(item, kind, [
      "Počet nových stanic / zbrojnic nelze určit automaticky.",
      "Výsledek: ruční ověření stávajícího pokrytí podle požárního poplachového plánu."
    ]);
    content.append(calculation);
    const explanation = createElement("div", "calculation-explanation");
    explanation.append(createElement("span", "calculation-kicker", "SLOVNÍ POPIS"));
    explanation.append(createElement("p", "",
      "Pěší dosah ani modelový středový bod neurčuje potřebu nové stanice. Prověřte kategorii JPO, skutečné dojezdové časy, výjezduschopnost a požární poplachový plán Jihomoravského kraje. HZS (profesionální JPO I) a JSDH obce jsou odlišné jednotky; případnou investici lze zadat až po doloženém závěru."));
    content.append(explanation);
    section.append(content);
    return section;
  }

  const housingCalculation = renderHousingContributionCalculation(item, kind);
  if (housingCalculation) {
    content.append(housingCalculation);
    section.append(content);
    return section;
  }
  if (isResidentParkingItem(item)) {
    const legal = getResidentParkingLegalEstimate(item);
    const recommended = getResidentParkingRecommendedEstimate(item);
    if (item.id === RESIDENT_PARKING_RESERVE_ID) {
      const reserveControls = renderResidentParkingReserveControls(item);
      if (reserveControls) content.append(reserveControls);
    }
    const calculation = createElement("div", "calculation-numeric");
    if (item.id === RESIDENT_PARKING_ID) {
      calculation.append(createElement("span", "calculation-kicker", "VÝPOČET DLE VYHLÁŠKY Č. 146/2024 SB."));
      const baseEquation = legal.requiredStalls === null
        ? "Zákonný počet nelze vypočítat: doplňte počet jednotek, HPP nebo ruční zákonnou plochu."
        : `${formatNumber(legal.floorArea)} m² celkem; běžná plocha ${formatNumber(legal.floorArea - legal.affordableRentalArea)} m² ÷ 120 + dostupná plocha ${formatNumber(legal.affordableRentalArea)} m² ÷ 240, násobeno korekcí ${formatNumber(legal.minimumFactor)} = ${formatNumber(legal.baseStalls)}; zaokrouhlení 0,5 nahoru = ${formatNumber(legal.requiredStalls)} stání.`;
      calculation.append(createElement("code", "calculation-equation",
        baseEquation));
      calculation.append(createElement("code", "calculation-equation",
        `Výsledek první parkovací varianty: ${formatNumber(legal.requiredStalls)} stání.`));
      content.append(calculation);
      content.append(createElement("p", "calculation-explanation",
        "Plocha je ručním přepisem dle právní definice nebo aproximací z HPP a průměrných ploch. Zvýhodněnou sazbu dostupného nájemního bydlení použijte jen při splnění podmínek zákona č. 283/2021 Sb."));
    } else if (item.id === RESIDENT_PARKING_LAW_ID) {
      calculation.append(createElement("span", "calculation-kicker", "ROZDÍL PRO SPLNĚNÍ ZÁKONA Č. 283/2021 SB."));
      calculation.append(createElement("code", "calculation-equation",
        recommended === null
          ? "Rozdíl nelze vypočítat bez počtu jednotek a zákonného základu."
          : `Modelová potřeba = ${formatNumber(recommended.apartmentUnits)} bytů × ${formatNumber(recommended.apartmentRate)} stání/byt + ${formatNumber(recommended.familyHomeUnits)} rodinných domů × ${formatNumber(recommended.familyHomeRate)} stání/dům = ${formatNumber(recommended.modeledStalls)} stání.`));
      calculation.append(createElement("code", "calculation-equation",
        recommended === null
          ? "Doplňte vstupy pro výpočet rozdílu."
          : `Doplňková druhá varianta = max(0, modelová potřeba − první varianta) = max(0, ${formatNumber(recommended.modeledStalls)} − ${formatNumber(legal.requiredStalls)}) = ${formatNumber(recommended.additionalStalls)} stání.`));
      content.append(calculation);
      content.append(createElement("p", "calculation-explanation",
        "Při realizaci obou variant se jejich počty sečtou bez duplicit: první varianta pokrývá vyhlášku č. 146/2024 Sb.; tato položka obsahuje pouze kladný zbytek. Je-li modelová potřeba nižší než první varianta, výsledek druhé je 0 stání."));
    } else {
      calculation.append(createElement("span", "calculation-kicker", "REZERVA V PROCENTECH NAD RÁMEC ZÁKONA Č. 283/2021 SB."));
      calculation.append(createElement("code", "calculation-equation",
        recommended === null
          ? "Rezervu nelze spočítat bez zákonného minima a modelové potřeby."
          : `Základ po započtení první a druhé varianty = ${formatNumber(recommended.combinedBaseStalls)} stání; rezerva = zaokrouhlení nahoru (${formatNumber(recommended.combinedBaseStalls)} × ${formatNumber(recommended.reservePercent * 100)} %) = ${formatNumber(recommended.reserveStalls)} stání.`));
      content.append(calculation);
      content.append(createElement("p", "calculation-explanation",
        "Rezerva se realizuje a vykazuje jako samostatná položka v procentech nad rámec zákona č. 283/2021 Sb.; není zahrnuta do počtu první ani druhé varianty. Procento upravíte tlačítkem „Upravit výpočet“."));
    }
    section.append(content);
    return section;
  }
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  if (projectEstimate?.type === "public-space" || item.id === "VYB_10B_1") {
    const calculation = renderProjectEstimateCalculation(item, kind);
    if (calculation) content.append(calculation);
    if (item.id === "VYB_10B_1") {
      content.append(createElement("p", "calculation-explanation",
        "Jde o orientační scénářovou projekci; metodika doporučuje prověřit existující sál a jeho volnou kapacitu, nikoli automaticky stavět nový."));
    }
    section.append(content);
    return section;
  }
  if (item.id === "VYB_14") {
    const calculation = renderProjectEstimateCalculation(item, kind);
    if (calculation) content.append(calculation);
    content.append(createElement("p", "calculation-explanation",
      "Pravidlo MOB-MET-001 doporučuje ověřit pěší dostupnost po skutečné trase ke středu nástupiště: typy zástavby A/B 500 m (300 m v souvislé kompaktní zástavbě), typy C/D 600–700 m. Modelové 1–2 zastávky nejsou zákonnou kvótou ani samostatně vypočteným počtem; jejich potřebu ověřte vůči stávající síti a linkám pro veřejnost."));
    const algorithms = renderCalculationAlgorithms(item);
    if (algorithms) content.append(algorithms);
    section.append(content);
    return section;
  }
  if (projectEstimate?.type === "infrastructure") {
    const calculation = renderProjectEstimateCalculation(item, kind);
    if (calculation) content.append(calculation);
    content.append(createElement("p", "calculation-explanation",
      "Jde o scénářovou kapacitní a investiční aproximaci. Nejde o samostatný požadavek na výstavbu, vyjádření správce sítě, RTS položku ani projektový rozpočet."));
    section.append(content);
    return section;
  }
  const polyclinicEstimate = projectEstimate;
  if (polyclinicEstimate?.type === "polyclinic") {
    const calculation = renderProjectEstimateCalculation(item, kind);
    if (calculation) content.append(calculation);
    const algorithms = renderCalculationAlgorithms(item);
    if (algorithms) content.append(algorithms);
    content.append(createElement("p", "calculation-explanation",
      `Jde o poměrný plošný a investiční scénář, nikoli požadavek na výstavbu. Nulový počet v původním excelovém modelu znamená pouze to, že hranice ${formatNumber(polyclinicEstimate.peoplePerCentre)} obyvatel na centrum nebyla dosažena.`));
    section.append(content);
    return section;
  }

  const equations = [];
  const sources = [];
  const comparison = getStandardComparison(item);
  const isStandardDelta = kind === "recommended" && getPrimary("REZIM_STD") !== "STANDARD"
    && comparison.delta !== null && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id);

  if (item.demand !== null) {
    const demandCalculation = formatWorkbookCalculation(
      item.sourceFormulas?.demand, getDemandLabel(item), item.demand, getDemandUnit(item)
    );
    equations.push(demandCalculation || `${getDemandLabel(item)} = ${formatNumber(item.demand)} ${text(getDemandUnit(item), "")}.`);
    const citation = calculationSource("Kapacitní poptávka", item.sourceFormulas?.demand);
    if (citation) sources.push(citation);
  }
  if (kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
    const percent = getSchoolRecommendationPercent(item);
    const baseCapacity = getStandardBaseQuantity(item);
    if (baseCapacity !== null) {
      const exactReserve = baseCapacity * percent / 100;
      const reserve = Math.ceil(exactReserve);
      equations.push(`Doporučená rezerva kapacity = navrhovaná kapacita × doporučené procento = ${formatNumber(baseCapacity)} × ${formatNumber(percent)} % = ${formatNumber(exactReserve)}; po zaokrouhlení nahoru ${formatNumber(reserve)} ${text(getStandardComparisonResultUnit(item), "jednotek")}.`);
      equations.push(`Kapacita včetně doporučené rezervy = navrhovaná kapacita + rezerva = ${formatNumber(baseCapacity)} + ${formatNumber(reserve)} = ${formatNumber(baseCapacity + reserve)} ${text(getStandardComparisonResultUnit(item), "jednotek")}.`);
    } else {
      equations.push(`Doporučené navýšení kapacity podle typu zástavby = ${formatNumber(percent)} %.`);
    }
  }
  if ((hasAppliedStandardTarget(item) || isStandardDelta) && comparison.delta !== null && !calc.hasOverride) {
    equations.push(`Základní kapacita = ${formatStandardComparisonAmount(item, comparison.base)}.`);
    equations.push(`Doporučené navýšení = ${formatStandardComparisonAmount(item, comparison.uplift)}.`);
    if (comparison.recommended !== null) {
      equations.push(`Cílová kapacita včetně navýšení = ${formatStandardComparisonAmount(item, comparison.base)} + ${formatStandardComparisonAmount(item, comparison.uplift)} = ${formatStandardComparisonAmount(item, comparison.recommended)}.`);
    }
    const verifiedCapacity = getVerifiedCapacity(item);
    const usableReserve = getAllowedCapacityUse(item, verifiedCapacity, comparison.uplift);
    equations.push(`Kapacita okolí započtená do doporučení = menší z doporučeného navýšení a ověřené kapacity × limit využití = min(${formatStandardComparisonAmount(item, comparison.uplift)}, ${formatStandardComparisonAmount(item, verifiedCapacity)} × ${formatNumber(getReserveLimit() * 100)} %) = ${formatStandardComparisonAmount(item, comparison.appliedCapacity)}.`);
    if (isStandardDelta) {
      equations.push(`Navýšení, které je třeba pokrýt = doporučené navýšení − započtená kapacita okolí = ${formatStandardComparisonAmount(item, comparison.uplift)} − ${formatStandardComparisonAmount(item, comparison.appliedCapacity)} = ${formatStandardComparisonAmount(item, comparison.delta)}.`);
    } else {
      equations.push(`Zbývající kapacita po započtení okolí = ${formatStandardComparisonAmount(item, calc.remainingCapacity ?? comparison.recommended)}.`);
    }
    if (usableReserve === 0 && verifiedCapacity > 0) {
      equations.push("Ověřená kapacita okolí se do doporučení nezapočítala, protože limit využití rezerv je nulový.");
    }
  }

  if (item.capacityCount !== null) {
    const calculation = formatWorkbookCalculation(
      item.sourceFormulas?.capacityCount, "Kapacitní potřeba", item.capacityCount, modelCountUnit(item)
    );
    equations.push(calculation || `Kapacitní potřeba = ${formatNumber(item.capacityCount)} ${text(modelCountUnit(item), "")}.`);
    const citation = calculationSource("Kapacitní potřeba", item.sourceFormulas?.capacityCount);
    if (citation) sources.push(citation);
  }
  if (item.spatialCount !== null) {
    const calculation = formatWorkbookCalculation(
      item.sourceFormulas?.spatialCount, "Prostorová potřeba", item.spatialCount, modelCountUnit(item)
    );
    equations.push(calculation || `Prostorová potřeba = ${formatNumber(item.spatialCount)} ${text(modelCountUnit(item), "")}.`);
    const citation = calculationSource("Prostorová potřeba", item.sourceFormulas?.spatialCount);
    if (citation) sources.push(citation);
  }
  if (item.capacityCount !== null && item.spatialCount !== null && item.modelProposed !== null) {
    equations.push(`Počet zařízení v návrhu = větší z kapacitní a prostorové potřeby = max(${formatNumber(item.capacityCount)}, ${formatNumber(item.spatialCount)}) = ${formatNumber(item.modelProposed)} ${text(modelCountUnit(item), "")}.`);
  } else if (item.modelProposed !== null) {
    equations.push(`Návrhový počet převzatý z modelu = ${formatNumber(item.modelProposed)} ${text(modelCountUnit(item), "")}.`);
  } else {
    equations.push("Zdrojový sešit neobsahuje číselný výsledný návrh.");
  }
  content.append(renderProjectEstimateCalculation(item, kind, equations, sources));
  const algorithms = renderCalculationAlgorithms(item);
  if (algorithms) content.append(algorithms);
  const explanation = createElement("div", "calculation-explanation");
  explanation.append(createElement("span", "calculation-kicker", "SLOVNÍ POPIS"));
  explanation.append(createElement("p", "",
    item.capacityCount !== null && item.spatialCount !== null
      ? `Výsledný počet zařízení se určí porovnáním kapacitní potřeby (${formatNumber(item.capacityCount)}) a prostorové potřeby (${formatNumber(item.spatialCount)}); použije se vyšší hodnota. Kapacitní část předchází plošným a investičním mezikrokům ve společném výpočtu.`
      : "Kapacitní část výpočtu předchází plošným a investičním mezikrokům ve společném bloku."));
  content.append(explanation);
  section.append(content);
  return section;
}

function renderCalculationAlgorithms(item) {
  const rules = [...new Map([...item.legalRules, ...item.recommendedRules, ...item.otherRules]
    .map((rule) => [String(rule[RULE_HEADERS.id]), rule])).values()]
    .filter((rule) => rule[RULE_HEADERS.formula] && rule[RULE_HEADERS.formula] !== "—");
  if (!rules.length) return null;
  const section = createElement("section", "calculation-methods");
  section.append(createElement("h4", "", "Metodické a podmínkové výpočty"));
  rules.forEach((rule) => {
    const block = createElement("article", "calculation-method");
    block.append(createElement("strong", "", `${displayRuleName(rule) || rule[RULE_HEADERS.area]} · ${rule[RULE_HEADERS.id]}`));
    const inputs = safeWorkbookDisplayText(readableWorkbookCitation(rule[RULE_HEADERS.inputs]));
    if (inputs) block.append(renderInputTokens(inputs));
    const formula = String(rule[RULE_HEADERS.formula]);
    const hasTechnicalFormula = isWorkbookImplementationText(formula);
    if (hasTechnicalFormula) {
      const description = safeWorkbookDisplayText(readableWorkbookCitation(rule[RULE_HEADERS.description]));
      if (description) block.append(createElement("p", "calculation-method-description", description));
    } else {
      block.append(renderAlgorithm(readableWorkbookCitation(formula)));
    }
    const output = safeWorkbookDisplayText(readableWorkbookCitation(rule[RULE_HEADERS.output]));
    if (output) {
      addField(block, "Výstup a jednotka", output);
    }
    const citation = [readableWorkbookCitation(rule[RULE_HEADERS.source]),
      readableWorkbookCitation(rule[RULE_HEADERS.location])]
      .filter((value) => value && !isWorkbookImplementationText(value)).join(" · ");
    if (citation) addField(block, "Zdroj a místo", citation, "citation-field");
    section.append(block);
  });
  return section;
}

function renderRulesDisclosure(item, kind, rules, displayName) {
  const section = createElement("details", "item-section rules-section");
  const isResidentParkingLegalSource = item.id === RESIDENT_PARKING_ID && kind === "legal";
  addDisclosureSummary(section, "Pravidla, výpočty a zdroje",
    isResidentParkingLegalSource ? "zákonný výpočet" : `${rules.length} pravidel`);
  const content = createElement("div", "rules-content");
  const cleanSourceName = cleanDisplayName(item.name);
  if (displayName !== cleanSourceName && !PARKING_SHARE_ASSUMPTIONS[item.id]) {
    content.append(createElement("p", "original-name-note", `Název ve zdrojovém sešitu: ${cleanSourceName}`));
  }
  const ruleBadges = createElement("div", "rule-badge-row");
  const isCombinedStandard = kind === "recommended" && getPrimary("REZIM_STD") === "STANDARD"
    && hasStandardComparison(item);
  const applicability = isResidentParkingLegalSource ? "Zákonný výpočet"
    : isCombinedStandard ? "Právní základ + doporučený standard"
      : kind === "legal" ? "Právní požadavek / kontrola"
      : getPrimary("REZIM_STD") === "STANDARD" ? "Ve zvoleném standardu" : "Volitelné doporučení";
  ruleBadges.append(createElement("span", `applicability-badge${kind === "recommended" && state.primaryValues.REZIM_STD !== "STANDARD" ? " is-optional" : ""}`, applicability));
  rules.forEach((rule) => ruleBadges.append(createElement("span", "rule-id", text(rule[RULE_HEADERS.id]))));
  content.append(ruleBadges);
  if (isResidentParkingLegalSource) {
    content.append(createElement("p", "item-results-note",
      "Vyhláška č. 146/2024 Sb., příloha č. 1 stanoví základní poměry 1/120 m² (běžné bydlení) a 1/240 m² (příslušné dostupné nájemní bydlení). Pro použití 1/240 m² ověřte splnění zákona č. 283/2021 Sb., § 13 písm. q); městský odkup sám o sobě nestačí. Podlahovou plochu dle právní definice a místní korekci ověřte v projektu. Rezerva i doporučený scénář jsou samostatně uvedeny v části Doporučené."));
  }
  rules.forEach((rule) => content.append(renderRuleDisclosure(rule, kind)));
  section.append(content);
  return section;
}

function renderRuleDisclosure(rule, kind) {
  const disclosure = createElement("details", "rule-disclosure");
  const heading = displayRuleName(rule) !== text(rule[RULE_HEADERS.equipment], "Pravidlo")
    ? displayRuleName(rule) : text(rule[RULE_HEADERS.area], "Pravidlo");
  addDisclosureSummary(disclosure, heading, text(rule[RULE_HEADERS.id], ""));
  disclosure.append(renderRule(rule, kind));
  return disclosure;
}

function renderItemCard(item, kind) {
  const legalRules = item.legalRules.filter((rule) =>
    !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id])) || item.id === "VYB_17B");
  const isCombinedRecommendation = kind === "recommended"
    && legalRules.length > 0 && item.recommendedRules.length > 0;
  const rules = kind === "legal" ? legalRules
    : isCombinedRecommendation || kind === "recommended" && getPrimary("REZIM_STD") === "STANDARD"
      ? [...new Map([...legalRules, ...item.recommendedRules]
        .map((rule) => [String(rule[RULE_HEADERS.id]), rule])).values()]
      : kind === "recommended" ? RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
        ? item.otherRules : item.recommendedRules : item.otherRules;
  const index = categoryIndex(item.category);
  const calc = adjustedResult(item, kind);
  const isSchoolRecommendation = kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
  const isDelta = kind === "recommended" && getPrimary("REZIM_STD") !== "STANDARD"
    && hasStandardComparison(item) && !isSchoolRecommendation;
  const comparison = isDelta ? getStandardComparison(item) : null;
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  const isTargetedEstimate = projectEstimate?.type === "targeted";
  const housingEstimate = getHousingContributionEstimate(item);
  const isPolyclinicEstimate = projectEstimate?.type === "polyclinic";
  const isInfrastructureEstimate = projectEstimate?.type === "infrastructure";
  const isPublicSpaceEstimate = projectEstimate?.type === "public-space";
  const isCulturalHallEstimate = item.id === "VYB_10B_1" && projectEstimate?.type === "item";
  const isZero = isTargetedEstimate ? false
    : housingEstimate ? housingEstimate.municipalApartments === 0
    : isPolyclinicEstimate ? projectEstimate.centerEquivalent === 0
    : isInfrastructureEstimate ? false
    : isPublicSpaceEstimate ? false
    : isCulturalHallEstimate ? projectEstimate.quantity === 0
    : isDelta ? comparison.delta === 0 : isRequirementSatisfied(calc);
  const realizationPlan = getRealizationPlan();
  const isSelectedForRealization = realizationPlan.selectedKeys.has(item.key);
  const details = createElement("details",
    `requirement-card category-${index % CATEGORY_COLORS.length} shape-${categoryShape(item.category)}${isZero ? " is-zero" : ""}${isSelectedForRealization ? " is-realization-selected" : ""}`);
  details.dataset.itemKey = item.key;
  details.dataset.kind = kind;
  const summary = document.createElement("summary");
  const top = createElement("div", "requirement-card-top");
  const realizationButton = createElement("button",
    `category-dot-control${isSelectedForRealization ? " is-selected" : ""}`);
  realizationButton.type = "button";
  realizationButton.dataset.realizationItemKey = item.key;
  realizationButton.setAttribute("aria-pressed", String(isSelectedForRealization));
  realizationButton.setAttribute("aria-label",
    `${isSelectedForRealization ? "Zrušit výběr" : "Vybrat"} k realizaci: ${displayCatalogItemName(item)}`);
  realizationButton.title = isSelectedForRealization ? "Odebrat z realizace" : "Vybrat k realizaci";
  realizationButton.append(createElement("span", `category-dot shape-${categoryShape(item.category)}`));
  top.append(realizationButton);
  const title = createElement("span", "requirement-title");
  const displayName = kind === "review" ? getReviewAlternativeName(item) : displayItemName(item);
  title.append(createElement("strong", "", displayName));
  if (isResidentParkingItem(item)) {
    title.append(createElement("small", "parking-part-label",
      item.id === RESIDENT_PARKING_ID ? "Počet dle vyhlášky" :
        item.id === RESIDENT_PARKING_LAW_ID ? "Pouze nezapočtený rozdíl" : "Vlastní vypočtený přírůstek"));
  }
  if (item.userAdded) {
    const userMark = createElement("small", "user-inserted-mark",
      "Vloženo uživatelem — bez zákonného podkladu");
    userMark.title = "Položka byla vložena uživatelem; není doložena jako zákonný podklad.";
    title.append(userMark);
  }
  if (isDelta || isCombinedRecommendation) {
    title.append(createElement("small", "", item.id === RESIDENT_PARKING_ID
      ? "Celková alternativa · nesčítá se s právním minimem"
      : "Rezerva a navýšení kapacity"));
  }
  top.append(title);
  const targetedCost = isTargetedEstimate
    ? projectEstimate.metrics.find((metric) => metric.label.includes("investice"))?.value
    : null;
  const summaryQuantity = createElement("span", "requirement-quantity", calc.hasOverride
    ? `${formatNumber(calc.required, 2)} ${text(getRequirementUnit(item), "")}`
    : isTargetedEstimate
    ? getTargetedPlanSummaryText(item, projectEstimate)
    : item.manualCoverageCheck
    ? "ověřit pokrytí"
    : housingEstimate ? `${formatNumber(housingEstimate.municipalApartments)} bytů pro město`
    : isPolyclinicEstimate ? `${formatNumber(projectEstimate.centerEquivalent, 3)} ekv. centra`
    : projectEstimate?.type === "infrastructure" ? `${formatNumber(projectEstimate.quantity)} ${projectEstimate.unit}`
    : isPublicSpaceEstimate ? `${formatNumber(projectEstimate.minimumArea)} m² minima`
    : isCulturalHallEstimate ? projectEstimate.quantity > 0 ? `${formatNumber(projectEstimate.quantity)} sál` : "prověřit podle populace"
    : isSchoolRecommendation ? `${formatNumber(getSchoolRecommendationPercent(item))} % doporučení`
    : isDelta ? comparison.delta === null ? "navýšení nezadáno"
      : `${formatStandardComparisonAmount(item, comparison.delta)} navíc`
      : item.id === "VYB_05B" ? getItemEstimateQuantity(item) > 0
        ? `${formatNumber(getItemEstimateQuantity(item))} ${getItemEstimateQuantity(item) === 1 ? "hala" : "haly"}`
        : "bez samostatné haly"
      : getSummaryRequirementQuantity(calc) === null ? "bez počtu"
        : `${formatNumber(getSummaryRequirementQuantity(calc), 2)} ${text(getRequirementUnit(item), "")}`);
  const summaryMetrics = createElement("span", "requirement-summary-metrics");
  summaryMetrics.append(summaryQuantity);
  if (calc.usesDemandBasedCapacity && !isDelta) {
    const capacitySummary = createElement("small", "requirement-capacity", getCapacitySummaryText(item, calc));
    capacitySummary.hidden = !calc.isCapacityVerified;
    summaryMetrics.append(capacitySummary);
  }
  summaryMetrics.append(createElement("small", "requirement-investment",
    item.manualCoverageCheck || isPublicSpaceEstimate
      || isTargetedEstimate && (targetedCost === null || targetedCost === undefined)
      ? "bez samostatného ocenění"
      : isTargetedEstimate ? formatCost(targetedCost)
      : projectEstimate?.cost !== null && projectEstimate?.cost !== undefined
        ? formatCost(projectEstimate.cost)
        : formatCost(getDisplayedCost(item, calc, kind))));
  top.append(summaryMetrics);
  top.append(createElement("span", "expand-mark", "＋"));
  summary.append(top);
  details.append(summary);

  const body = createElement("div", "requirement-body");
  const parkingLegalInputs = renderResidentParkingLegalInputs(item, kind);
  body.append(isInfrastructureEstimate
    ? renderInfrastructureResults(item, kind)
    : renderItemMetrics(item, kind, isDelta));
  const healthEditor = renderHealthCheckEditor(item);
  if (isTargetedEstimate) {
    const inlineHealthCheck = Boolean(healthEditor && isHealthcareItem(item));
    if (inlineHealthCheck) {
      const verification = renderTargetedVerificationSection(item, kind);
      if (verification) healthEditor.querySelector(".health-check-editor")?.append(verification);
    }
  }
  body.append(renderItemCalculation(item, kind));
  const assumptions = renderObjectAssumptionControls(item, kind);
  if (assumptions) body.append(assumptions);
  const availabilitySection = renderCapacityEditor(item, kind);
  if (parkingLegalInputs) availabilitySection.append(parkingLegalInputs);
  if (isDelta || isSchoolRecommendation
    || (kind === "recommended" && getPrimary("REZIM_STD") === "STANDARD" && hasStandardComparison(item))) {
    availabilitySection.append(renderStandardComparison(item));
  }
  if (item.quantityBasis && !isResidentParkingItem(item)) {
    availabilitySection.append(createElement("p", "quantity-basis-note",
      `Počet je převzat z „${MODEL_HEADERS.demand}“ (${formatNumber(item.demand)} ${text(getDemandUnit(item), "")}); hodnota N_navrh v modelu byla ${formatNumber(item.sourceModelQuantity)} a představovala počet zařízení/bodů, nikoli kapacitu uvedenou v jednotce položky.`));
  }
  if (healthEditor) availabilitySection.append(healthEditor);
  if (kind === "recommended") {
    const spatialEstimate = renderSpatialEstimate(item);
    if (spatialEstimate) {
      const spatialDisclosure = createElement("details", "item-section spatial-section");
      addDisclosureSummary(spatialDisclosure, "Prostorový odhad dostupnosti");
      spatialDisclosure.append(spatialEstimate);
      availabilitySection.append(spatialDisclosure);
    }
  }
  const alternativeNote = renderAlternativeContext(item);
  if (alternativeNote) availabilitySection.append(alternativeNote);
  body.append(availabilitySection);
  if (rules.length || item.id === RESIDENT_PARKING_ID && kind === "legal") {
    body.append(renderRulesDisclosure(item, kind, rules, displayName));
  }
  const transferredWarnings = renderTransferredReviewWarnings(item, rules);
  if (transferredWarnings) body.append(transferredWarnings);
  details.append(body);
  return details;
}

function renderStandardComparison(item) {
  const saved = state.standardComparisons[item.key] || { legal: "", recommended: "", source: "" };
  const result = getStandardComparison(item);
  const section = createElement("section", "standard-comparison");
  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
    section.append(createElement("h4", "", "Doporučené navýšení kapacity"));
    section.append(createElement("p", "standard-comparison-note",
      "Metodické doporučení se zadává v procentech, nikoli jako zákonné minimum v počtu dětí nebo žáků. Výchozí procento odpovídá typu zástavby: A/B 10 %, C 15 %, D 20 %."));
    const controls = createElement("div", "standard-comparison-fields");
    const wrapper = createElement("label", "capacity-field");
    wrapper.append(createElement("span", "capacity-field-label", "Doporučené navýšení kapacity (%)"));
    const input = document.createElement("input");
    input.type = "number";
    input.min = "0";
    input.max = "100";
    input.step = "any";
    input.value = String(getSchoolRecommendationPercent(item));
    input.placeholder = "Doporučené navýšení v %";
    input.dataset.standardItem = item.key;
    input.dataset.standardField = "recommended";
    wrapper.append(input);
    controls.append(wrapper);
    section.append(controls);
    section.append(createElement("p", "standard-comparison-result",
      `Doporučené navýšení kapacity: ${formatNumber(getSchoolRecommendationPercent(item))} %.`));
    return section;
  }
  section.append(createElement("h4", "", "Doporučené navýšení kapacity"));
  section.append(createElement("p", "standard-comparison-note",
    `Zadejte přímo doporučené navýšení v jednotce vybavenosti (${text(getStandardComparisonUnit(item), "jednotka položky")}). Nejde o zákonnou hodnotu ani procento. Do navýšení lze započítat ověřenou kapacitu okolí v téže jednotce.`));
  const controls = createElement("div", "standard-comparison-fields");
  controls.append(createStandardComparisonInput(item, "recommended",
    `Doporučené navýšení (${text(getStandardComparisonUnit(item), "MJ")})`, saved.recommended));
  section.append(controls);
  const resultText = result.delta === null
    ? `Zadejte doporučené navýšení v ${text(getStandardComparisonUnit(item), "jednotce vybavenosti")}.`
    : `Doporučené navýšení ${formatStandardComparisonAmount(item, result.uplift)}; započtená kapacita okolí ${formatStandardComparisonAmount(item, result.appliedCapacity)}; zbývá pokrýt ${formatStandardComparisonAmount(item, result.delta)}; orientační cena navýšení ${formatCost(result.cost)}.`;
  section.append(createElement("p", "standard-comparison-result", resultText));
  return section;
}

function createStandardComparisonInput(item, field, label, value) {
  const wrapper = createElement("label", "capacity-field");
  wrapper.append(createElement("span", "capacity-field-label", label));
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.step = isWholeNumberUnit(getStandardComparisonUnit(item)) ? "1" : "any";
  input.value = value ?? "";
  input.placeholder = `Navýšení v ${text(getStandardComparisonUnit(item), "jednotce položky")}`;
  input.dataset.standardItem = item.key;
  input.dataset.standardField = field;
  wrapper.append(input);
  return wrapper;
}

function updateStandardComparison(event) {
  const input = event.target.closest("[data-standard-item][data-standard-field]");
  if (!input) return;
  const item = state.items.find((candidate) => candidate.key === input.dataset.standardItem);
  if (!item) return;
  const field = input.dataset.standardField;
  {
    const value = numericOrNull(input.value);
    if (input.value !== "" && (value === null || value < 0
      || (field === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id) && value > 100)
      || (field === "recommended" && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
        && isWholeNumberUnit(getStandardComparisonUnit(item)) && !Number.isInteger(value)))) {
      input.setCustomValidity(field === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
        ? "Zadejte doporučené navýšení mezi 0 a 100 %."
        : isWholeNumberUnit(getStandardComparisonUnit(item))
          ? "Zadejte nezáporný celý počet jednotek."
          : "Zadejte nezápornou hodnotu.");
      if (event.type === "change") input.reportValidity();
      return;
    }
    input.setCustomValidity("");
  }
  const saved = state.standardComparisons[item.key] || { legal: "", recommended: "", source: "" };
  state.standardComparisons[item.key] = { ...saved, [field]: input.value };
  saveScenario();
  updateStandardComparisonCard(item);
  if (event.type === "change") sortRequirementCards(elements.recommendedList);
  updateScenarioStatus();
}

function updateStandardComparisonCard(item) {
  const result = getStandardComparison(item);
  document.querySelectorAll(`.requirement-card[data-item-key="${CSS.escape(item.key)}"]`).forEach((card) => {
    const kind = card.dataset.kind || (card.closest("#recommended-list") ? "recommended" : "legal");
    const isSchoolRecommendation = kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
    const isDelta = kind === "recommended" && getPrimary("REZIM_STD") !== "STANDARD"
      && hasStandardComparison(item) && !isSchoolRecommendation;
    const calc = adjustedResult(item, kind);
    const resultLine = card.querySelector(".standard-comparison-result");
    if (resultLine) {
      resultLine.textContent = isSchoolRecommendation
        ? `Doporučené navýšení kapacity: ${formatNumber(getSchoolRecommendationPercent(item))} %.`
        : result.delta === null
          ? `Zadejte doporučené navýšení v ${text(getStandardComparisonUnit(item), "jednotce vybavenosti")}.`
          : `Doporučené navýšení ${formatStandardComparisonAmount(item, result.uplift)}; započtená kapacita okolí ${formatStandardComparisonAmount(item, result.appliedCapacity)}; zbývá pokrýt ${formatStandardComparisonAmount(item, result.delta)}; orientační cena navýšení ${formatCost(result.cost)}.`;
    }
    if (getPrimary("REZIM_STD") !== "STANDARD" && !card.querySelector(".requirement-title small")) {
      card.querySelector(".requirement-title")?.append(createElement("small", "", "Rezerva a navýšení kapacity"));
    }
    const quantity = card.querySelector(".requirement-quantity");
    const investment = card.querySelector(".requirement-investment");
    const projectEstimate = getProjectPlanningEstimate(item, kind);
    const displayedInvestment = projectEstimate?.cost ?? getDisplayedCost(item, calc, kind);
    if (isSchoolRecommendation) {
      if (quantity) quantity.textContent = `${formatNumber(getSchoolRecommendationPercent(item))} % doporučení`;
      if (investment) investment.textContent = formatCost(displayedInvestment);
    } else if (isDelta) {
      if (quantity) quantity.textContent = result.delta === null
        ? "navýšení nezadáno" : `${formatStandardComparisonAmount(item, result.delta)} navíc`;
      if (investment) investment.textContent = formatCost(displayedInvestment);
      card.classList.toggle("is-zero", result.delta === 0);
    } else {
      updateRequirementCardSummary(card, item, calc);
      if (investment) investment.textContent = formatCost(displayedInvestment);
      card.classList.toggle("is-zero", isRequirementSatisfied(calc));
    }
    const capacityResult = card.querySelector(".capacity-result");
    if (capacityResult) updateCapacityResultText(capacityResult, item, calc, kind);
    const metrics = card.querySelector(".item-results");
    if (metrics) metrics.replaceWith(renderItemMetrics(item, kind, isDelta));
    const calculation = card.querySelector(".calculation-section");
    if (calculation) {
      const replacement = renderItemCalculation(item, kind);
      calculation.replaceChildren(...replacement.childNodes);
    }
  });

}

function calculateSpatialEstimate(item) {
  const config = SPATIAL_ESTIMATE_CONFIG[item.id];
  if (!config) return null;
  const settlementType = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  const distanceMeters = config.distanceMeters === "by-settlement-type"
    ? item.id === "VYB_03C" ? settlementType === "D" ? 1000 : 500
      : settlementType === "D" ? null : 300
    : config.distanceMeters;
  const rule = item.recommendedRules.find((candidate) => candidate[RULE_HEADERS.id] === config.ruleId);
  if (distanceMeters === null) {
    return {
      rule,
      distanceMeters: null,
      note: "Metodika pro typ D nestanoví pevnou vzdálenost; je nutné ověřit veřejnou parkovou zeleň v každém obytném sídle."
    };
  }
  const areaCode = getSecondary("Plocha pro odhad dostupnosti") || "S_zastav";
  const areaHa = numericOrNull(getPrimary(areaCode));
  const shapeCode = getPrimary("KOD_TVAR");
  const shape = GEOMETRY_EXTENTS[shapeCode];
  const overlap = numericOrNull(getSecondary(COVERAGE_OVERLAP_INPUT.code));
  const detour = numericOrNull(state.pedestrianDetourFactor);
  if (areaHa === null || areaHa <= 0 || !shape || overlap === null || !detour || detour <= 0) {
    return { rule, distanceMeters, note: "Doplňte platnou plochu, geometrický tvar a koeficient pěší sítě." };
  }
  const boundingArea = areaHa * 10000 * shape.areaFactor;
  const longSide = Math.sqrt(boundingArea * shape.aspect);
  const shortSide = Math.sqrt(boundingArea / shape.aspect);
  const radius = distanceMeters / detour;
  const spacing = Math.SQRT2 * radius * (1 - overlap);
  const columns = Math.max(1, Math.ceil(longSide / spacing));
  const rows = Math.max(1, Math.ceil(shortSide / spacing));
  return {
    rule,
    count: columns * rows,
    columns,
    rows,
    areaHa,
    areaCode,
    shapeCode,
    distanceMeters,
    overlap,
    detour,
    radius,
    spacing,
    longSide,
    shortSide
  };
}

function displayShapeName(code) {
  const field = PRIMARY_INPUTS.find((candidate) => candidate.code === "KOD_TVAR");
  const index = field?.options.indexOf(code) ?? -1;
  return index < 0 ? code : field.labels[index];
}

function renderSpatialEstimate(item) {
  const config = SPATIAL_ESTIMATE_CONFIG[item.id];
  if (!config) return null;
  const estimate = calculateSpatialEstimate(item);
  const section = createElement("section", "spatial-estimate");
  section.append(createElement("h4", "", "Orientační počet podle plošné sítě kruhových dosahů"));
  const source = estimate.rule
    ? [estimate.rule[RULE_HEADERS.source], estimate.rule[RULE_HEADERS.location]].filter(Boolean).join(" · ")
    : config.ruleId;
  if (estimate.count === undefined) {
    section.append(createElement("p", "spatial-estimate-note", estimate.note));
    section.append(createElement("p", "spatial-estimate-source", `Pravidlo: ${source}`));
    return section;
  }
  const result = createElement("p", "spatial-estimate-result");
  result.append(createElement("strong", "", `${estimate.count} ${text(item.unit, "zařízení")}`));
  result.append(createElement("span", "", `Mřížka ${estimate.columns} × ${estimate.rows} bodů`));
  section.append(result);
  const formula = createElement("code", "spatial-estimate-formula",
    `R = ${formatNumber(estimate.distanceMeters)} / ${formatNumber(estimate.detour)} = ${formatNumber(estimate.radius)} m; d = √2 × R × (1 − ${formatNumber(estimate.overlap * 100)} %) = ${formatNumber(estimate.spacing)} m; N = ceil(L/d) × ceil(W/d)`);
  section.append(formula);
  section.append(createElement("p", "spatial-estimate-note",
    `Vychází z plochy ${formatNumber(estimate.areaHa)} ha (${estimate.areaCode}), tvaru ${displayShapeName(estimate.shapeCode)} a dosahu ${formatNumber(estimate.distanceMeters)} m po pěší síti. Kružnice jsou převedeny na vzdušný poloměr koeficientem uliční sítě ${formatNumber(estimate.detour)}; překrytí ${formatNumber(estimate.overlap * 100)} % zkracuje rozteč mřížky, nejde o doslovný podíl překryté plochy kruhů.`));
  section.append(createElement("p", "spatial-estimate-note",
    "Konzervativní odhad pokrytí obalového obdélníku zvoleného tvaru při rovnoměrném rozmístění. Neumisťuje body do skutečné uliční sítě ani neověřuje bariéry, vstupy, existující zařízení či kapacitu; polohy musí potvrdit GIS. Doporučený standard není zákonná povinnost."));
  if (source) section.append(createElement("p", "spatial-estimate-source", `Metodický podklad: ${source}`));
  const button = createElement("button", "text-button spatial-estimate-action", "Přenést jako uživatelský návrh");
  button.type = "button";
  button.dataset.spatialEstimate = item.key;
  section.append(button);
  return section;
}

function useSpatialEstimate(event) {
  const button = event.target.closest("[data-spatial-estimate]");
  if (!button) return;
  const item = state.items.find((candidate) => candidate.key === button.dataset.spatialEstimate);
  if (!item) return;
  const estimate = calculateSpatialEstimate(item);
  if (!estimate || !Number.isFinite(estimate.count)) return;
  const card = button.closest(".requirement-card");
  const quantity = card?.querySelector('[data-capacity-field="override"]');
  const note = card?.querySelector('[data-capacity-field="overrideNote"]');
  if (!quantity || !note) return;
  quantity.value = String(estimate.count);
  note.value = `Plošný odhad (${estimate.rule?.[RULE_HEADERS.id] || "bez metodického pravidla"}): ${estimate.count} bodů; mřížka ${estimate.columns} × ${estimate.rows}, dosah ${estimate.distanceMeters} m, překrytí ${formatNumber(estimate.overlap * 100)} %. Nutné potvrdit GIS.`;
  updateCapacityEdit({ target: quantity });
  updateCapacityEdit({ target: note });
  showNotice(`Plošný odhad ${estimate.count} ${text(item.unit, "zařízení")} byl uložen jako uživatelský návrh. Ověřte skutečné polohy v GIS.`);
}

function createTextInput(label, value, dataset, placeholder = "") {
  const wrapper = createElement("label", "capacity-field");
  wrapper.append(createElement("span", "capacity-field-label", label));
  const isMultiline = dataset.field === "note" || dataset.field === "overrideNote";
  const input = document.createElement(isMultiline ? "textarea" : "input");
  if (isMultiline) input.rows = 2;
  else input.type = dataset.field === "date" ? "date" : "text";
  input.value = value ?? "";
  input.placeholder = placeholder;
  input.dataset.recordId = dataset.id;
  input.dataset.capacityField = dataset.field;
  if (dataset.kind) input.dataset.capacityKind = dataset.kind;
  wrapper.append(input);
  return wrapper;
}

function renderCapacityEditor(item, kind) {
  const disclosure = createElement("details", "item-section capacity-section");
  addDisclosureSummary(disclosure, "Ověřený stav v okolí a uživatelský návrh");
  disclosure.append(renderAvailabilitySearch(item));

  const editor = createElement("section", "capacity-editor");
  const heading = createElement("div", "capacity-heading");
  heading.append(createElement("span", "capacity-unit", `Jednotka pro ověření: ${text(getCapacityUnit(item))}`));
  editor.append(heading);

  const calc = adjustedResult(item, kind);
  const controls = createElement("div", "capacity-controls");
  const capacity = getCapacityEdit(item);
  const capacityUnit = getCapacityUnit(item);
  const capLabel = capacityUnit
    ? `Celková ověřená kapacita v okolí (${capacityUnit})`
    : "Celková ověřená kapacita v okolí";
  const capWrap = createTextInput(capLabel, capacity.quantity, { id: item.key, field: "quantity", kind: "number" }, "Prázdné = neověřeno");
  const capInput = capWrap.querySelector("input");
  capInput.type = "number";
  capInput.min = "0";
  capInput.step = "any";
  capInput.setAttribute("aria-label", `${capLabel} pro ${displayItemName(item)}`);
  controls.append(capWrap);
  const guidance = createElement("p", "capacity-search-guide", capacitySearchGuide(item));
  editor.append(guidance);
  controls.append(createTextInput("Zdroj / ověření (nepovinné)", capacity.source, { id: item.key, field: "source" }, "Např. GIS, plán nebo správce"));
  controls.append(createTextInput("Datum (doplní se automaticky)", capacity.date, { id: item.key, field: "date" }, "Vyplní se při zadání kapacity"));
  controls.append(createTextInput("Poznámka k ověření", capacity.note, { id: item.key, field: "note" }, "Rozsah, kapacita a omezení"));
  const parkingDemand = renderParkingDemand(item);
  if (parkingDemand) editor.append(parkingDemand);

  const override = state.overrides[item.key] || { quantity: "", note: "" };
  const overrideWrap = createTextInput("Uživatelsky upravený návrh", override.quantity, { id: item.key, field: "override", kind: "number" }, "Ponechat hodnotu modelu");
  const overrideInput = overrideWrap.querySelector("input");
  overrideInput.type = "number";
  overrideInput.min = "0";
  overrideInput.step = "any";
  overrideInput.setAttribute("aria-label", `Uživatelsky upravený návrh pro ${displayItemName(item)}`);
  controls.append(overrideWrap);
  controls.append(createTextInput("Poznámka k úpravě návrhu", override.note, { id: item.key, field: "overrideNote" }, "Povinné pro použití přepisu"));
  editor.append(controls);

  const result = createElement("div", "capacity-result");
  result.dataset.resultFor = item.key;
  updateCapacityResultText(result, item, calc, kind);
  editor.append(result);
  const caution = createElement("p", "capacity-caution",
    "Zadanou kapacitu lze použít bez vyplnění zdroje; datum se nastaví automaticky. Bez zdroje jde o uživatelský údaj bez nezávislého doložení. Ruční návrh se do výsledku promítne pouze po doplnění poznámky. U škol a veřejných prostranství se porovnává v kapacitní jednotce (žáci/děti nebo m²), nikoli v počtu budov/souborů.");
  editor.append(caution);
  disclosure.append(editor);
  return disclosure;
}

function renderResidentParkingLegalInputs(item, kind) {
  if (item.id !== RESIDENT_PARKING_ID || kind !== "legal") return null;
  const legal = getResidentParkingLegalEstimate(item);
  const section = createElement("section", "parking-legal-inputs");
  section.append(createElement("h4", "", "Vstupy a odhad zákonné plochy"));
  section.append(createElement("p", "parking-legal-note",
    "Pracovní odhad se skládá z počtu bytových jednotek, typového podílu bytových domů, průměrných ploch a očekávaného podílu bytů nad 120 m². HPP ani tento modelový odhad nenahrazují zákonnou podlahovou plochu dle § 13 písm. n) zákona č. 283/2021 Sb.; znáte-li ji, zadejte ruční přepis."));
  const fields = createElement("div", "parking-legal-fields");
  const estimateSummary = createElement("div", "parking-legal-estimate");
  estimateSummary.append(createElement("strong", "",
    legal.estimatedFloorArea === null ? "Odhad plochy nelze určit" :
      `Modelově odhadnutá podlahová plocha: ${formatNumber(legal.estimatedFloorArea)} m²`));
  if (legal.units !== null) {
    estimateSummary.append(createElement("span", "",
      `${formatNumber(legal.units)} jednotek · ${formatNumber(legal.apartmentUnits)} bytů v bytových domech · ${formatNumber(legal.familyHomeUnits)} rodinných domů`));
  }
  if (legal.estimatedFloorArea !== null) {
    estimateSummary.append(createElement("span", "",
      `Plocha podle zadaných průměrů: ${formatNumber(legal.smallApartmentUnits)} bytů × ${formatNumber(legal.smallApartmentArea)} m² + ${formatNumber(legal.largeApartmentUnits)} bytů × ${formatNumber(legal.largeApartmentArea)} m² + ${formatNumber(legal.familyHomeUnits)} rodinných domů × ${formatNumber(legal.familyHomeAreaPerUnit)} m². Z městského podílu se při předpokladu 1b zohlední ${formatNumber(legal.accessibleShare * 100)} % plochy bytových domů.`));
  }
  fields.append(estimateSummary);
  const area = createTextInput("Ruční přepis plochy dle právní definice (nepovinné)",
    legal.manualFloorArea, { id: item.key, field: "parkingLegalArea" }, "m²");
  const areaInput = area.querySelector("input");
  areaInput.type = "number";
  areaInput.min = "0";
  areaInput.step = "any";
  fields.append(area);
  const factor = createTextInput("Minimální místní korekce", legal.minimumFactor,
    { id: item.key, field: "parkingLegalFactor" }, "Výchozí 1, pokud není stanovena");
  const factorInput = factor.querySelector("input");
  factorInput.type = "number";
  factorInput.min = "0";
  factorInput.step = "any";
  fields.append(factor);
  section.append(fields);
  if (legal.floorArea === null) {
    section.append(createElement("p", "result-note result-warning",
      "Doplňte počet bytových jednotek, celkovou HPP, nebo ruční zákonnou plochu, aby šel výpočet provést."));
  }
  return section;
}

function renderResidentParkingReserveControls(item) {
  if (item.id !== RESIDENT_PARKING_RESERVE_ID) return null;
  const disclosure = createElement("details", "project-estimate-controls parking-reserve-controls");
  addDisclosureSummary(disclosure, "Upravit výpočet");
  disclosure.append(createElement("p", "project-estimate-input-note",
    "Zadejte procento navíc z celkového základu obou parkovacích variant. Tato rezerva se přičítá jako samostatná položka."));
  const fields = createElement("div", "secondary-fields project-estimate-inputs");
  const field = {
    code: "PARKING_RESERVE_PERCENT",
    label: "Rezerva nad rámec zákona č. 283/2021 Sb.",
    kind: "percent",
    min: 0,
    max: 1,
    step: 0.01,
    unit: "%"
  };
  fields.append(createInputControl(field, getResidentParkingReservePercent(), true,
    `${item.id}-recommended`, "itemEstimate", item.key));
  disclosure.append(fields);
  return disclosure;
}

function capacitySearchGuide(item) {
  const place = String(getPrimary("OBEC") || "zadané lokality").trim();
  if (item.id === "VYB_04A") {
    return `V ${place} zjistěte součet ploch veřejných prostranství započitatelných do zákonného minima. Zadejte celkovou plochu v m²; komunikace se nezapočítávají. Doporučený zdroj ověření je územní/GIS podklad nebo návrh projektu.`;
  }
  if (DEMAND_LABELS[item.id]) {
    return `V ${place} a v relevantním spádu ověřte dostupný počet míst pro ${DEMAND_SEARCH_LABELS[item.id]}. Zadejte počet míst ve stejné jednotce, který může nový záměr skutečně využít. Doporučujeme zaznamenat zařízení a rozsah spádu.`;
  }
  return `V ${place} a v územním nebo dojezdovém spádu ověřte volnou kapacitu pro „${displayItemName(item)}“. Zadejte hodnotu v jednotce ${text(getCapacityUnit(item))}; doporučujeme uvést konkrétní zařízení/lokalitu a rozsah ověření.`;
}

function availabilitySearchUrl(item) {
  const place = String(getPrimary(PROJECT_LOCATION_INPUT.code) || PROJECT_LOCATION_INPUT.defaultValue).trim();
  const query = [
    displayItemName(item),
    place,
    "dostupnost v okolí, poskytovatelé a zařízení, volná kapacita, spádové území"
  ].join(" ");
  return `https://www.google.com/search?q=${encodeURIComponent(query)}`;
}

function renderAvailabilitySearch(item) {
  const section = createElement("div", "availability-search");
  const link = createElement("a", "text-button availability-search-link",
    "Hledat dostupnost vybavenosti v okolí ↗");
  link.href = availabilitySearchUrl(item);
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  section.append(link);
  section.append(createElement("p", "item-results-note",
    "Odkaz otevře webové vyhledávání podle názvu vybavenosti a lokality. Výsledky nejsou samy o sobě ověřením kapacity; zaznamenejte konkrétní zařízení a ověřený zdroj."));
  return section;
}

function updateCapacityResultText(target, item, calc, kind) {
  target.replaceChildren();
  const line = createElement("div", "result-line");
  const isSchoolRecommendation = kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
  const isDelta = kind === "recommended" && getPrimary("REZIM_STD") !== "STANDARD"
    && hasStandardComparison(item) && !isSchoolRecommendation;
  const comparison = isDelta ? getStandardComparison(item) : null;
  if (isSchoolRecommendation) {
    line.append(createElement("strong", "", `Doporučené navýšení kapacity: ${formatNumber(getSchoolRecommendationPercent(item))} %`));
  } else if (isDelta) {
    if (comparison.delta === null) {
      line.append(createElement("strong", "", "Doporučené navýšení nebylo zadáno."));
    } else {
      line.append(createElement("span", "", `Doporučený přírůstek: ${formatStandardComparisonAmount(item, comparison.uplift)}`));
      line.append(createElement("span", "", `Kapacita okolí započtená do limitu: ${formatStandardComparisonAmount(item, comparison.appliedCapacity)}`));
      line.append(createElement("strong", "", `Zbývá pokrýt: ${formatStandardComparisonAmount(item, comparison.delta)}`));
    }
  } else if (item.id === RESIDENT_PARKING_ID && kind === "legal") {
    line.append(createElement("span", "", calc.required === null
      ? "Zákonné minimum: chybí podlahová plocha podle právní definice."
      : `Zákonné minimum: ${formatValue(calc.required, item.unit)}`));
    if (calc.required !== null) {
      line.append(createElement("span", "", `Ověřená kapacita v okolí: ${formatValue(calc.availableCapacity, item.unit)}`));
      line.append(createElement("strong", "", `Zbývá pokrýt: ${formatValue(calc.remaining, item.unit)}`));
    }
  } else if (calc.usesDemandBasedCapacity) {
    const resultUnit = getCapacityResultUnit(item);
    line.append(createElement("span", "", `Kapacitní potřeba: ${formatValue(calc.capacityDemand, resultUnit)}`));
    line.append(createElement("span", "", `Ověřená kapacita v okolí: ${formatValue(calc.availableCapacity, resultUnit)}`));
    line.append(createElement("span", "", `Započteno do výpočtu: ${formatValue(calc.verifiedCapacity, resultUnit)}`));
    line.append(createElement("strong", "", `Zbývá pokrýt: ${formatValue(calc.remainingCapacity, resultUnit)}`));
  } else {
    line.append(createElement("span", "", `Návrh: ${formatValue(calc.required, item.unit)}`));
    line.append(createElement("span", "", `Ověřená kapacita v okolí: ${formatValue(calc.availableCapacity, item.unit)}`));
    line.append(createElement("span", "", `Započteno do výpočtu: ${formatValue(calc.verifiedCapacity, item.unit)}`));
    line.append(createElement("strong", "", `Zbývá: ${formatValue(calc.remaining, item.unit)}`));
  }
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  const displayedCost = projectEstimate?.cost ?? getDisplayedCost(item, calc, kind);
  line.append(createElement("strong", "result-investment", formatCost(displayedCost)));
  target.append(line);
  if (item.id === RESIDENT_PARKING_ID && kind === "legal" && calc.required === null) {
    target.append(createElement("p", "result-note",
      "Doplňte počet bytových jednotek, celkovou HPP, nebo ruční přepis podlahové plochy dle právní definice."));
  } else if (!isDelta && !isSchoolRecommendation && calc.required === null) {
    target.append(createElement("p", "result-note", "Excel pro tuto položku nemá číselný výsledný návrh; cenu proto nelze bezpečně přepočítat."));
  } else if (displayedCost === null) {
    target.append(createElement("p", "result-note", "V Excelu není číselná celková cena a množství, ze kterých by šla odvodit jednotková cena."));
  } else if (calc.hasOverride) {
    target.append(createElement("p", "result-note", "Uživatelský návrh je aktivní. Jde o váš projektový přepis, nikoli změnu právního nebo metodického zdroje."));
  } else if (calc.isCapacityVerified) {
    target.append(createElement("p", "result-note", "Ověřená kapacita se odečetla od potřeby ve stejné jednotce. Přepočet ceny je poměrný scénář podle součtu v Excelu, nikoli nová kalkulace RTS."));
    if (!getCapacityEdit(item).source?.trim()) {
      target.append(createElement("p", "result-note result-warning", "Kapacita je započtena jako uživatelský údaj, ale není doložena zdrojem."));
    }
  }
  if (calc.isCapacityVerified || isDelta && comparison?.delta !== null) {
    target.append(createElement("p", "result-note",
      `Započtení je omezeno nejvýše na ${formatNumber(getReserveLimit() * 100)} % zadané kapacity okolí a nepřekročí potřebu; u doporučeného standardu nesníží zákonné minimum.`));
  }
}

function renderRule(rule, kind) {
  const card = createElement("article", `rule-card ${kind}-rule-card`);
  const heading = createElement("div", "rule-heading");
  const tagClass = kind === "legal" ? "tag-legal" : kind === "recommended" ? "tag-method" : "tag-review";
  const ruleId = String(rule[RULE_HEADERS.id] ?? "").trim();
  const tagText = ruleId === "ZDR-LEG-001" || ruleId === "ZDR-SPEC-001" ? "Právní limit dostupnosti služby"
    : kind === "legal" ? "Zákonná povinnost"
    : kind === "recommended" ? "Doporučený standard" : "Další pravidlo";
  heading.append(createElement("span", `rule-tag ${tagClass}`, tagText));
  heading.append(createElement("code", "rule-id", text(rule[RULE_HEADERS.id])));
  card.append(heading);
  const displayName = displayRuleName(rule);
  const originalName = cleanDisplayName(text(rule[RULE_HEADERS.equipment], "Pravidlo"));
  card.append(createElement("h4", "", displayName !== originalName
    ? displayName : text(rule[RULE_HEADERS.area], "Pravidlo")));
  if (rule.appMapping) {
    const mapping = safeWorkbookDisplayText(
      `${text(rule[RULE_HEADERS.equipmentId])} → ${rule.appLinkedIds.join(", ")} (sémantické přiřazení podle ID pravidla; původní sešit zůstává beze změny)`
    );
    if (mapping) addField(card, "Vazba v aplikaci", mapping);
  }
  if (rule[RULE_HEADERS.type]) addField(card, "Druh pravidla", rule[RULE_HEADERS.type]);
  const description = safeWorkbookDisplayText(rule[RULE_HEADERS.description]);
  const condition = safeWorkbookDisplayText(rule[RULE_HEADERS.condition]);
  const inputs = safeWorkbookDisplayText(rule[RULE_HEADERS.inputs]);
  const formula = safeWorkbookDisplayText(rule[RULE_HEADERS.formula]);
  const output = safeWorkbookDisplayText(rule[RULE_HEADERS.output]);
  const rounding = safeWorkbookDisplayText(rule[RULE_HEADERS.rounding]);
  const verification = safeWorkbookDisplayText(rule[RULE_HEADERS.verification]);
  const limitation = safeWorkbookDisplayText(rule[RULE_HEADERS.limitation]);
  if (description) addField(card, "Co pravidlo stanoví", description);
  if (condition) addField(card, "Podmínka / rozsah", condition);
  if (inputs) card.append(renderInputTokens(inputs));
  if (formula && formula !== "—") card.append(renderAlgorithm(formula));
  if (output) addField(card, "Výstup a jednotka", output);
  if (rounding) addField(card, "Zaokrouhlení", rounding);
  if (verification) addField(card, "Stav ověření", verification);
  if (limitation) addField(card, "Omezení / nutné rozhodnutí", limitation);
  const citation = [rule[RULE_HEADERS.source], rule[RULE_HEADERS.location]]
    .map(readableWorkbookCitation)
    .filter((value) => value && !isWorkbookImplementationText(value)).join(" · ");
  if (citation) addField(card, "Zdroj a místo v metodice", citation, "citation-field");
  appendSourceLink(card, rule);
  if (requiresCurrentVerification(rule)) card.append(renderManualReview(rule));
  const searchLinks = renderCurrentVerificationLinks(rule);
  if (searchLinks) card.append(searchLinks);
  return card;
}

function requiresCurrentVerification(rule) {
  const id = String(rule[RULE_HEADERS.id] ?? "").trim();
  const checkText = `${rule[RULE_HEADERS.limitation] ?? ""} ${rule[RULE_HEADERS.description] ?? ""}`;
  return ["ZDR-LEG-001", "ZDR-SPEC-001", "CULT-COMM-001", "SAFE-MET-001",
    "SAFE-LEG-001", "ZEL-LEG-002", "TECH-LEGAL-002"].includes(id)
    || /konkrétn(?:í|e) povinnosti.*aktuáln(?:í|ím) práv(?:u|ě).*krajsk(?:ém|ého) plánu/i.test(checkText);
}

function renderCurrentVerificationLinks(rule) {
  const checkText = `${rule[RULE_HEADERS.limitation] ?? ""} ${rule[RULE_HEADERS.description] ?? ""}`;
  if (!requiresCurrentVerification(rule)
    && !/aktuáln(?:í|ím) práv(?:u|ě).*krajsk(?:ém|ého) plánu/i.test(checkText)) return null;
  const place = String(getPrimary(PROJECT_LOCATION_INPUT.code) || PROJECT_LOCATION_INPUT.defaultValue).trim();
  const subject = String(displayRuleName(rule) || rule[RULE_HEADERS.area] || "").trim();
  const retailQuery = String(rule[RULE_HEADERS.id] ?? "").trim() === "CULT-COMM-001";
  const links = retailQuery
    ? [
      ["Celostátní předpisy a standardy", `maloobchodní vybavenost obecný právní poměr m2 prodejní plochy na obyvatele Česká republika`],
      ["Územní plán města Brna", `Územní plán města Brna ${place} maloobchod komerční občanská vybavenost regulativy`],
      ["Regulační podmínky lokality", `${place} maloobchod územní plán regulační plán podmínky využití ploch`]
    ]
    : [
      ["Aktuální právní předpis", `${subject} ${place} aktuální právní předpis`],
      ["Krajský plán – ZÚR Jihomoravského kraje", `Zásady územního rozvoje Jihomoravského kraje ${place} ${subject}`],
      ["Místní územní plán", `Územní plán města Brna ${place} ${subject}`]
    ];
  const section = createElement("section", "current-verification");
  section.append(createElement("h4", "", "Dohledat aktuální podmínky pro zadanou lokalitu"));
  addField(section, "Lokalita", place);
  const inputs = createElement("div", "detail-field");
  inputs.append(createElement("span", "field-label", "Údaje potřebné při ověření"));
  inputs.append(renderInputTokens(rule[RULE_HEADERS.inputs] || "Viz podmínku a zdroj pravidla"));
  section.append(inputs);
  const projectContext = createElement("div", "current-project-context");
  projectContext.append(createElement("strong", "", `Lokalita: ${place}`));
  ["O_celk", "O_sidlo", "O_spad", "TYP_UZEMI", "S_uzemi", "S_zastav"].forEach((code) => {
    const value = getPrimary(code);
    if (value === null || value === undefined || value === "") return;
    const descriptor = inputDescriptor(code);
    const label = state.primaryLabels[code] || descriptor[INPUT_HEADERS.name];
    projectContext.append(createElement("span", "",
      `${text(label, code)}: ${formatValue(value, text(descriptor[INPUT_HEADERS.unit], ""))}`));
  });
  const population = derivedValues();
  projectContext.append(createElement("span", "",
    `Celkem po záměru ve spádové oblasti: ${formatValue(population.O_prah, "obyvatel")}`));
  section.append(projectContext);
  const linkList = createElement("div", "current-verification-links");
  links.forEach(([label, query]) => {
    const link = createElement("a", "source-link", label);
    link.href = `https://www.google.com/search?q=${encodeURIComponent(query)}`;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    linkList.append(link);
  });
  if (retailQuery) {
    const officialPlan = createElement("a", "source-link", "Oficiální portál Územního plánu města Brna ↗");
    officialPlan.href = "https://upmb.brno.cz/";
    officialPlan.target = "_blank";
    officialPlan.rel = "noopener noreferrer";
    linkList.append(officialPlan);
  }
  section.append(linkList);
  section.append(createElement("p", "current-verification-note",
    "Odkazy spouštějí vyhledávání, nejsou potvrzením závaznosti. Rozhodující je aktuální znění předpisu a územně-plánovací dokumentace pro konkrétní záměr."));
  return section;
}

function renderInputTokens(value) {
  const wrapper = createElement("div", "visual-field");
  wrapper.append(createElement("span", "field-label", "Potřebné vstupy"));
  const chips = createElement("div", "input-token-list");
  const tokens = String(value).split(/[;\n]+/).map((part) => part.trim()).filter(Boolean);
  (tokens.length ? tokens : [String(value)]).forEach((token) => chips.append(createElement("span", "input-token", token)));
  wrapper.append(chips);
  return wrapper;
}

function splitAlgorithmBranches(value) {
  const source = String(value).trim();
  if (!/\bIF\s*\(/i.test(source)) return null;
  const chunks = source.split(/;\s*/).filter(Boolean);
  return chunks.map((chunk, index) => {
    const trimmed = chunk.trim();
    const isElse = /^ELSE\b/i.test(trimmed);
    const isFallback = /^(?:ELSE\s+)?(?:JINAK|OTHERWISE)\b/i.test(trimmed);
    const expression = trimmed.replace(/^ELSE\s*/i, "");
    const match = expression.match(/^IF\s*\(([\s\S]*)\)\s*$/i);
    if (!match) {
      return { label: isFallback ? "JINAK" : `VĚTEV ${index + 1}`, condition: "", result: trimmed };
    }
    const inner = match[1];
    let depth = 0;
    let commaIndex = -1;
    for (let i = 0; i < inner.length; i += 1) {
      if (inner[i] === "(") depth += 1;
      else if (inner[i] === ")") depth -= 1;
      else if (inner[i] === "," && depth === 0) {
        commaIndex = i;
        break;
      }
    }
    return {
      label: isElse ? `JINAK, POKUD ${index + 1}` : `POKUD ${index + 1}`,
      condition: commaIndex < 0 ? inner : inner.slice(0, commaIndex).trim(),
      result: commaIndex < 0 ? "" : inner.slice(commaIndex + 1).trim()
    };
  });
}

function evaluateConditionClause(condition) {
  const match = condition.match(/^(?:TYP_UZEMI\s*=\s*)?([A-D](?:\s*\/\s*[A-D])*)(?:\s+AND\s+(.+))?$/i);
  if (!match) return null;
  const selectedType = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  const allowedTypes = match[1].split("/").map((type) => type.trim().toUpperCase());
  if (!selectedType || !allowedTypes.includes(selectedType)) return false;
  if (!match[2]) return true;

  const comparison = match[2].match(/^(O_sidlo|O_celk|O_prah|S_uzemi|S_zastav|rho_hruba|rho_cista)\s*(>=|<=|>|<|==|=)\s*(-?\d+(?:[.,]\d+)?)$/i);
  if (!comparison) return null;
  const variable = comparison[1].toLowerCase();
  const left = variable === "o_sidlo" || variable === "o_prah" || variable.startsWith("rho_")
    ? derivedValues()[comparison[1]]
    : numericOrNull(getPrimary(comparison[1]));
  const right = Number(comparison[3].replace(",", "."));
  if (left === null || !Number.isFinite(right)) return null;
  switch (comparison[2]) {
    case ">": return left > right;
    case "<": return left < right;
    case ">=": return left >= right;
    case "<=": return left <= right;
    case "=":
    case "==": return left === right;
    default: return null;
  }
}

function evaluateAlgorithmBranches(branches) {
  let decided = false;
  let unresolvedEarlier = false;
  for (const branch of branches) {
    if (!branch.condition) {
      branch.status = decided ? "inactive" : unresolvedEarlier ? "unknown" : "fallback";
      if (!unresolvedEarlier) decided = true;
      continue;
    }
    if (decided) {
      branch.status = "inactive";
      continue;
    }
    const result = evaluateConditionClause(branch.condition);
    if (result === null) {
      branch.status = "unknown";
      unresolvedEarlier = true;
    } else if (result && !unresolvedEarlier) {
      branch.status = "active";
      decided = true;
    } else if (result) {
      branch.status = "unknown";
    } else {
      branch.status = unresolvedEarlier ? "unknown" : "inactive";
    }
  }
}

function renderAlgorithm(value) {
  const wrapper = createElement("div", "algorithm-visual");
  wrapper.append(createElement("span", "field-label", "Vzorec / algoritmus"));
  const branches = splitAlgorithmBranches(value);
  if (!branches) {
    const formula = createElement("pre", "algorithm-single");
    formula.textContent = String(value);
    wrapper.append(formula);
    return wrapper;
  }
  evaluateAlgorithmBranches(branches);
  const flow = createElement("div", "algorithm-branches");
  branches.forEach((branch, index) => {
    if (index > 0) flow.append(createElement("span", "algorithm-or", "NEBO"));
    const card = createElement("div", `algorithm-branch${branch.status ? ` is-${branch.status}` : ""}`);
    card.append(createElement("span", "branch-label", branch.label));
    if (branch.status) {
      const statusText = branch.status === "active" ? "Použije se"
        : branch.status === "inactive" ? "Nesplněno / předchozí větev"
          : branch.status === "fallback" ? "Výchozí větev" : "Nelze jednoznačně vyhodnotit";
      card.append(createElement("span", "branch-status", statusText));
    }
    if (branch.condition) {
      card.append(createElement("span", "branch-caption", "VSTUP / PODMÍNKA"));
      card.append(createElement("code", "branch-condition", branch.condition));
    }
    if (branch.result) {
      card.append(createElement("span", "branch-caption", "VÝSLEDEK"));
      card.append(createElement("strong", "branch-result", branch.result));
    }
    flow.append(card);
  });
  wrapper.append(flow);
  return wrapper;
}

function addField(parent, label, value, className = "") {
  const wrapper = createElement("div", `detail-field ${className}`.trim());
  wrapper.append(createElement("span", "field-label", label));
  wrapper.append(createElement("div", "field-value", text(value)));
  parent.append(wrapper);
}

function appendSourceLink(parent, rule) {
  const rawUrl = rule[RULE_HEADERS.url];
  if (typeof rawUrl !== "string" || !rawUrl.trim()) return;
  let url;
  try {
    url = new URL(rawUrl.trim());
  } catch {
    return;
  }
  if (url.protocol !== "https:" && url.protocol !== "http:") return;
  const link = createElement("a", "source-link", "Otevřít zdroj ↗");
  link.href = url.href;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  parent.append(link);
}

function renderRequirements() {
  const matching = matchingItems();
  const selectedKeys = getRealizationPlan().selectedKeys;
  const visibleCount = state.hideUnrealizedItems
    ? matching.filter((item) => selectedKeys.has(item.key)).length
    : matching.length;
  const selectedItem = state.items.find((item) => item.key === state.selectedItemKey);
  const categoryName = selectedItem ? displayItemName(selectedItem)
    : state.selectedCategory || "Všechny skupiny vybavenosti";
  elements.selectedCategoryTitle.textContent = categoryName;
  elements.visibleCount.textContent = state.hideUnrealizedItems
    ? `${visibleCount} realizovaných`
    : `${visibleCount} položek`;
  elements.hideUnrealizedItems.checked = state.hideUnrealizedItems;
  renderRequirementList(elements.legalList, "legal", matching);
  renderRequirementList(elements.infrastructureList, "infrastructure", matching);
  renderRequirementList(elements.recommendedList, "recommended", matching);
  renderRequirementList(elements.reviewList, "review", matching);
  for (const [name, list] of [
    ["legal", elements.legalList],
    ["infrastructure", elements.infrastructureList],
    ["recommended", elements.recommendedList],
    ["review", elements.reviewList]
  ]) {
    const collapsed = Boolean(state.collapsedLists[name]);
    list.hidden = collapsed;
    const button = elements.requirementColumns.querySelector(`[data-collapse-list="${name}"]`);
    if (!button) continue;
    button.setAttribute("aria-expanded", String(!collapsed));
    const sectionName = button.closest(".requirement-column").querySelector("h3").textContent;
    button.setAttribute("aria-label", `${collapsed ? "Rozbalit" : "Sbalit"} ${sectionName}`);
  }
}

function showNotice(message, isError = false) {
  elements.notice.textContent = message;
  elements.notice.classList.toggle("notice-error", isError);
  elements.notice.hidden = false;
}

function collectFeedbackNotes() {
  const notes = [];
  const add = (location, label, value, relatedValues = []) => {
    const content = String(value ?? "").trim();
    if (!content) return;
    const related = relatedValues
      .map(([name, relatedValue]) => [name, String(relatedValue ?? "").trim()])
      .filter(([, relatedValue]) => relatedValue);
    const details = related.length
      ? `\nSouvisející hodnoty: ${related.map(([name, relatedValue]) => `${name}: ${relatedValue}`).join("; ")}`
      : "";
    notes.push({ label: `${location} — ${label}`, content: `${content}${details}` });
  };
  const itemLocation = (item) =>
    `Položka: ${displayItemName(item)} [${item.id}] · skupina: ${item.category || "neuvedená"}`;
  const getItem = (key) => state.items.find((item) => item.key === key);

  for (const [key, edit] of Object.entries(state.capacityEdits)) {
    const item = getItem(key);
    const location = item ? itemLocation(item) : `Položka z uloženého scénáře [${key}]`;
    const related = [
      ["ověřená kapacita", edit.quantity],
      ["datum ověření", edit.date]
    ];
    add(location, "zdroj ověření kapacity", edit.source, related);
    add(location, "datum ověření kapacity", edit.date, [["ověřená kapacita", edit.quantity]]);
    add(location, "poznámka k ověření kapacity", edit.note, related);
  }
  for (const [key, override] of Object.entries(state.overrides)) {
    const item = getItem(key);
    const location = item ? itemLocation(item) : `Položka z uloženého scénáře [${key}]`;
    add(location, "poznámka k uživatelsky upravenému návrhu", override.note,
      [["uživatelsky upravený návrh", override.quantity]]);
  }
  for (const [key, checks] of Object.entries(state.healthChecks)) {
    const item = getItem(key);
    const itemLabel = item ? itemLocation(item) : `Zdravotnická položka z uloženého scénáře [${key}]`;
    for (const check of Array.isArray(checks) ? checks : []) {
      const location = `${itemLabel} · odbornost: ${check.profession || "neuvedena"}`;
      const related = [
        ["limit dojezdu (min)", check.limitMinutes],
        ["poskytovatel / pracoviště", check.provider],
        ["ověřený dojezd (min)", check.travelMinutes],
        ["ověřená kapacita (osob)", check.capacity],
        ["potřeba ze zdroje (osob)", check.demand]
      ];
      add(location, "zdroj / datum ověření", check.source, related);
      add(location, "poznámka", check.note, related);
    }
  }

  const rules = new Map();
  for (const item of state.items) {
    for (const rule of [...item.legalRules, ...item.recommendedRules, ...item.otherRules]) {
      const id = reviewKey(rule);
      if (!rules.has(id)) rules.set(id, { item, rule });
    }
  }
  for (const [id, check] of Object.entries(state.reviewChecks)) {
    const match = rules.get(id);
    const location = match
      ? `${itemLocation(match.item)} · kontrola: ${displayRuleName(match.rule)} [${match.rule[RULE_HEADERS.id]}]`
      : `Kontrola pravidla z uloženého scénáře [${id}]`;
    const related = [
      ["stav kontroly", check.status],
      ["zjištěný stav / výsledek", check.finding],
      ["zdroj ověření", check.source]
    ];
    add(location, "stav kontroly", check.status, related);
    add(location, "zdroj ověření", check.source, related);
    add(location, "zjištěný stav / výsledek", check.finding, related);
    add(location, "poznámka / další krok", check.note, related);
  }
  return notes;
}

function formatFeedbackReport(errors, changes) {
  const notes = collectFeedbackNotes();
  const selectedItem = state.items.find((item) => item.key === state.selectedItemKey);
  const location = selectedItem
    ? `${selectedItem.category} / ${displayItemName(selectedItem)} [${selectedItem.id}]`
    : state.selectedCategory || "Všechny skupiny vybavenosti";
  const sections = [
    "ZPĚTNÁ VAZBA K APLIKACI",
    `Vytvořeno: ${new Date().toLocaleString("cs-CZ")}`,
    `Načtený projekt / sešit: ${state.fileName || "neuveden"}`,
    `Místo v aplikaci při vytvoření: ${location}`,
    "",
    "POZOROVANÉ CHYBY",
    errors.trim() || "—",
    "",
    "NAVRHOVANÉ ZMĚNY",
    changes.trim() || "—",
    "",
    "POZNÁMKY A PODKLADY ZAPSANÉ U POLOŽEK A KONTROL",
    notes.length
      ? notes.map(({ label, content }) => `${label}:\n${content}`).join("\n\n")
      : "Žádné další poznámky."
  ];
  return `\uFEFF${sections.join("\r\n")}\r\n`;
}

function exportFeedbackReport(event) {
  event.preventDefault();
  const errors = document.querySelector("#feedback-errors").value;
  const changes = document.querySelector("#feedback-changes").value;
  if (!errors.trim() && !changes.trim() && !collectFeedbackNotes().length) {
    document.querySelector("#feedback-status").textContent =
      "Nejprve popište chybu, navrhovanou změnu nebo doplňte poznámku v aplikaci.";
    return;
  }

  const report = formatFeedbackReport(errors, changes);
  const timestamp = new Date();
  const datePart = [
    timestamp.getFullYear(),
    String(timestamp.getMonth() + 1).padStart(2, "0"),
    String(timestamp.getDate()).padStart(2, "0")
  ].join("-");
  const timePart = [
    String(timestamp.getHours()).padStart(2, "0"),
    String(timestamp.getMinutes()).padStart(2, "0")
  ].join("-");
  let objectUrl;
  try {
    const file = new Blob([report], { type: "text/plain;charset=utf-8" });
    const downloadLink = document.createElement("a");
    objectUrl = URL.createObjectURL(file);
    downloadLink.href = objectUrl;
    downloadLink.download = `zpetna-vazba-${datePart}_${timePart}.txt`;
    downloadLink.click();
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  } catch (error) {
    document.querySelector("#feedback-status").textContent =
      `Soubor .txt se nepodařilo vytvořit nebo stáhnout: ${error instanceof Error ? error.message : String(error)}`;
    if (objectUrl) URL.revokeObjectURL(objectUrl);
    return;
  }
  const status = document.querySelector("#feedback-status");
  status.textContent = "Soubor .txt byl stažen. V otevřeném e-mailu jej přiložte a zprávu odešlete.";
  const subject = encodeURIComponent("Zpětná vazba k aplikaci Vybavenost území");
  const body = encodeURIComponent(
    "Dobrý den,\n\nv příloze posílám export zpětné vazby k aplikaci.\n\n"
    + "Soubor .txt je potřeba přiložit ručně."
  );
  try {
    window.location.href = `mailto:xavavrusa@vutbr.cz?subject=${subject}&body=${body}`;
  } catch (error) {
    status.textContent =
      `Soubor .txt byl stažen, ale e-mailový klient se nepodařilo otevřít: ${error instanceof Error ? error.message : String(error)}`;
  }
}

function clearNotice() {
  elements.notice.hidden = true;
  elements.notice.textContent = "";
  elements.notice.classList.remove("notice-error");
}

function updateCapacityEdit(event) {
  const input = event.target.closest("[data-record-id][data-capacity-field]");
  if (!input) return;
  const item = state.items.find((candidate) => candidate.key === input.dataset.recordId);
  if (!item) return;
  const id = item.key;
  const field = input.dataset.capacityField;
  const numericFields = ["quantity", "override", "parkingShare", "parkingLegalArea", "parkingLegalFactor"];
  const numericValue = numericOrNull(input.value);
  const numericInvalid = numericFields.includes(field)
    && (input.validity.badInput || (input.value !== "" && (numericValue === null
      || numericValue < 0 || (field === "parkingShare" && numericValue > 100))));
  if (numericInvalid) {
    input.setCustomValidity(field === "parkingShare"
      ? "Zadejte podíl mezi 0 a 100 %."
      : "Zadejte nezápornou číselnou hodnotu nebo pole ponechte prázdné.");
    if (event.type === "change") input.reportValidity();
    return;
  }
  input.setCustomValidity("");
  if (field === "parkingShare") {
    const share = numericOrNull(input.value);
    state.parkingShares[id] = input.value === "" ? "" : share;
  } else if (field === "parkingLegalArea" || field === "parkingLegalFactor") {
    const saved = state.itemEstimateValues[item.key] || {};
    const code = field === "parkingLegalArea" ? "PARKING_LEGAL_AREA_M2" : "PARKING_LEGAL_FACTOR_MIN";
    state.itemEstimateValues[item.key] = { ...saved, [code]: input.value === "" ? "" : Number(input.value) };
  } else if (field === "quantity") {
    const existing = getCapacityEdit(item);
    const date = input.value !== "" && !existing.date ? getLocalIsoDate() : existing.date;
    state.capacityEdits[id] = { ...existing, quantity: input.value, date };
    if (date) {
      const dateInput = input.closest(".capacity-controls")?.querySelector('[data-capacity-field="date"]');
      if (dateInput instanceof HTMLInputElement) dateInput.value = date;
    }
  } else if (field === "override") {
    const existing = state.overrides[id] || { quantity: "", note: "" };
    state.overrides[id] = { ...existing, quantity: input.value === "" ? "" : Number(input.value) };
  } else if (field === "overrideNote") {
    const existing = state.overrides[id] || { quantity: "", note: "" };
    state.overrides[id] = { ...existing, note: input.value };
  } else {
    const existing = getCapacityEdit(item);
    state.capacityEdits[id] = { ...existing, [field]: input.value };
  }
  saveScenario();
  if ((field === "parkingLegalArea" || field === "parkingLegalFactor") && event.type === "change") {
    const disclosureState = captureRequirementDisclosureState();
    renderRequirements();
    restoreRequirementDisclosureState(disclosureState);
    updateScenarioStatus();
    return;
  }
  const cards = [...document.querySelectorAll(".requirement-card")]
    .filter((candidate) => candidate.dataset.itemKey === id);
  for (const card of cards) {
    const calc = adjustedResult(item, card.dataset.kind);
    const isDelta = card.dataset.kind === "recommended"
      && getPrimary("REZIM_STD") !== "STANDARD" && hasStandardComparison(item)
      && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
    const isSchoolRecommendation = card.dataset.kind === "recommended"
      && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
    const comparison = isDelta ? getStandardComparison(item) : null;
    card.classList.toggle("is-zero", isDelta
      ? comparison.delta === 0 : isRequirementSatisfied(calc));
    const title = card.querySelector(".requirement-title strong");
    if (title) title.textContent = displayItemName(item);
    const result = card.querySelector(".capacity-result");
    if (result) updateCapacityResultText(result, item, calc, card.dataset.kind);
    const calculation = card.querySelector(".calculation-section");
    if (calculation) {
      const replacement = renderItemCalculation(item, card.dataset.kind);
      calculation.replaceChildren(...replacement.childNodes);
    }
    const metrics = card.querySelector(".item-results");
    if (metrics) {
      metrics.replaceWith(renderItemMetrics(item, card.dataset.kind, isDelta));
    }
    if (isSchoolRecommendation) {
      const quantity = card.querySelector(".requirement-quantity");
      if (quantity) quantity.textContent = `${formatNumber(getSchoolRecommendationPercent(item))} % doporučení`;
    } else {
      updateRequirementCardSummary(card, item, calc);
    }
    const summaryInvestment = card.querySelector(".requirement-investment");
    const projectEstimate = getProjectPlanningEstimate(item, card.dataset.kind);
    if (summaryInvestment) summaryInvestment.textContent =
      formatCost(projectEstimate?.cost ?? getDisplayedCost(item, calc, card.dataset.kind));
    if (field === "parkingShare") {
      const current = card.querySelector(".parking-share-assumption");
      const replacement = renderParkingDemand(item);
      if (current && replacement) {
        current.querySelector(".parking-demand").textContent =
          replacement.querySelector(".parking-demand").textContent;
        current.querySelector(".capacity-caution").textContent =
          replacement.querySelector(".capacity-caution").textContent;
      }
      card.querySelectorAll('[data-capacity-field="parkingShare"]').forEach((shareInput) => {
        if (shareInput !== input) shareInput.value = getParkingShare(item) ?? "";
      });
    }
    if (card !== input.closest(".requirement-card")) {
      card.querySelectorAll("[data-capacity-field][data-record-id]").forEach((otherInput) => {
        const source = otherInput.dataset.capacityField === "override" || otherInput.dataset.capacityField === "overrideNote"
          ? state.overrides[id]
          : state.capacityEdits[id];
        const fieldKey = otherInput.dataset.capacityField === "override" ? "quantity"
          : otherInput.dataset.capacityField === "overrideNote" ? "note" : otherInput.dataset.capacityField;
        if (source && fieldKey in source) otherInput.value = source[fieldKey] ?? "";
      });
    }
  }
  if (hasStandardComparison(item)) updateStandardComparisonCard(item);
  if (event.type === "change") {
    for (const parent of [elements.legalList, elements.recommendedList,
      elements.infrastructureList, elements.reviewList]) {
      sortRequirementCards(parent);
    }
  }
  updateScenarioStatus();
}

function getLocalIsoDate() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function sortRequirementCards(parent) {
  const cards = [...parent.querySelectorAll(":scope > .requirement-card")];
  const sortCards = (list) => list.sort((left, right) =>
    Number(left.classList.contains("is-zero")) - Number(right.classList.contains("is-zero")));
  sortCards(cards);
  parent.querySelectorAll(".reserve-capacity-group").forEach((group) => {
    const groupedCards = [...group.querySelectorAll(":scope > .requirement-card")];
    sortCards(groupedCards);
    group.append(...groupedCards);
  });
  if (!cards.length) return;
  const unlinked = [...parent.querySelectorAll(":scope > .unlinked-rule-card")];
  parent.append(...cards, ...unlinked);
}

function updateHealthcareCheck(event) {
  const input = event.target.closest("[data-health-item][data-health-id][data-health-field]");
  if (!input) return;
  const item = state.items.find((candidate) => candidate.key === input.dataset.healthItem);
  if (!item) return;
  const field = input.dataset.healthField;
  if (["limitMinutes", "travelMinutes", "capacity", "demand"].includes(field)
    && (input.validity.badInput || (input.value !== ""
      && (numericOrNull(input.value) === null || Number(input.value) < 0)))) {
    input.setCustomValidity("Zadejte nezápornou číselnou hodnotu nebo pole ponechte prázdné.");
    if (event.type === "change") input.reportValidity();
    return;
  }
  input.setCustomValidity("");
  const checks = getHealthChecks(item);
  const entry = checks.find((record) => record.id === input.dataset.healthId);
  if (!entry) return;
  entry[field] = input.value;
  state.healthChecks[item.key] = checks;
  saveScenario();
  for (const editor of document.querySelectorAll(".health-check-editor")) {
    if (editor.dataset.healthEditor !== item.key) continue;
    const record = [...editor.querySelectorAll(".health-check-record")]
      .find((candidate) => candidate.dataset.healthRecord === entry.id);
    if (!record) continue;
    record.querySelectorAll("[data-health-field]").forEach((otherInput) => {
      if (otherInput !== input) otherInput.value = entry[otherInput.dataset.healthField] ?? "";
    });
    const title = record.querySelector(".health-check-top strong");
    if (title) title.textContent = entry.profession || "Nová odbornost";
    const link = record.querySelector(".health-search-link");
    if (link) link.href = healthSearchUrl(item, entry);
    const result = record.querySelector(".health-check-result");
    if (result) result.textContent = healthCapacitySummary(entry);
    const allLink = editor.querySelector(".health-search-all");
    if (allLink) allLink.href = healthSearchUrl(item);
  }
  updateScenarioStatus();
}

function addHealthcareCheck(event) {
  const button = event.target.closest("[data-add-health-item]");
  if (!button) return;
  const item = state.items.find((candidate) => candidate.key === button.dataset.addHealthItem);
  if (!item) return;
  const checks = getHealthChecks(item);
  const entry = {
    id: `added-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    profession: "",
    group: "",
    limitMinutes: "",
    provider: "",
    travelMinutes: "",
    capacity: "",
    demand: "",
    source: "",
    note: ""
  };
  checks.push(entry);
  state.healthChecks[item.key] = checks;
  saveScenario();
  for (const editor of document.querySelectorAll(".health-check-editor")) {
    if (editor.dataset.healthEditor === item.key) {
      editor.querySelector(".health-check-records")?.append(renderHealthCheckRecord(item, entry));
    }
  }
  updateScenarioStatus();
}

function handleRequirementClick(event) {
  useSpatialEstimate(event);
  addHealthcareCheck(event);
}

function updateManualReview(event) {
  const input = event.target.closest("[data-review-id][data-review-field]");
  if (!input) return;
  const id = input.dataset.reviewId;
  const existing = state.reviewChecks[id] || { status: "Nezahájeno", source: "", finding: "", note: "" };
  state.reviewChecks[id] = { ...existing, [input.dataset.reviewField]: input.value };
  saveScenario();
  updateScenarioStatus();
}

function createProjectExportSheet(scenario) {
  const rows = [
    ["Přenositelný projekt — Vybavenost území"],
    ["Formát", PROJECT_SHEET_MARKER],
    ["Zdrojový sešit", state.fileName],
    ["Exportováno", new Date().toISOString()],
    [],
    ["VSTUPY PROJEKTU"],
    ["Skupina", "Parametr", "Kód", "Hodnota"]
  ];
  const primaryFields = new Map(PRIMARY_INPUTS.map((field) => [field.code, field]));
  const secondaryFields = new Map([
    ...SECONDARY_INPUTS,
    ...PROJECT_ESTIMATE_INPUTS,
    ...HOUSING_CONTRIBUTION_INPUTS,
    ...ALTERNATIVE_ESTIMATE_INPUTS
  ].map((field) => [field.code, field]));
  const appendValues = (group, values, fields) => {
    for (const [code, value] of Object.entries(values)) {
      const field = fields.get(code);
      rows.push([group, field?.label || state.primaryLabels[code] || code, code, value ?? ""]);
    }
  };
  appendValues("Hlavní vstupy", scenario.primaryValues, primaryFields);
  appendValues("Další vstupy", scenario.secondaryValues, secondaryFields);
  rows.push([], ["VÝBĚR K REALIZACI"], ["Vybavenost", "ID", "Tematická skupina", "Realizováno"]);
  const selectedKeys = new Set(Object.entries(scenario.realizationSelections)
    .filter(([, selected]) => Number(selected) === 1).map(([key]) => key));
  state.items.forEach((item) => rows.push([
    displayCatalogItemName(item), item.id, item.category, selectedKeys.has(item.key) ? "Ano" : "Ne"
  ]));
  rows.push([], ["OVĚŘENÁ KAPACITA A UŽIVATELSKÝ NÁVRH"],
    ["Vybavenost", "ID", "Ověřená kapacita", "Zdroj", "Datum", "Poznámka k ověření", "Návrhová hodnota", "Poznámka k návrhu"]);
  state.items.forEach((item) => {
    const capacity = scenario.capacityEdits[item.key] || {};
    const override = scenario.overrides[item.key] || {};
    if (![capacity.quantity, capacity.source, capacity.date, capacity.note, override.quantity, override.note]
      .some((value) => value !== undefined && value !== "" && value !== null)) return;
    rows.push([
      displayCatalogItemName(item), item.id, capacity.quantity ?? "", capacity.source ?? "",
      capacity.date ?? "", capacity.note ?? "", override.quantity ?? "", override.note ?? ""
    ]);
  });
  rows.push([], ["NASTAVENÍ ZOBRAZENÍ"],
    ["Přepínač", "Hodnota"]);
  rows.push(["Skrýt nerealizované", scenario.hideUnrealizedItems ? "Ano" : "Ne"]);
  rows.push(["Sbalené tematické skupiny", scenario.collapsedCategories.join("; ")]);
  rows.push(["Sbalené výsledkové seznamy", Object.entries(scenario.collapsedLists)
    .filter(([, collapsed]) => collapsed).map(([name]) => name).join("; ")]);
  rows.push(["Vybraná skupina", scenario.selectedCategory]);
  rows.push(["Vyhledávací filtr", scenario.search]);
  rows.push(["Odvozené ukazatele rozbalené", scenario.derivedDetailsOpen ? "Ano" : "Ne"]);
  rows.push(["Barevný režim", scenario.theme]);
  rows.push([], ["DATA SCÉNÁŘE PRO OPĚTOVNÝ IMPORT", "Nemažte řádky označené JSON_PART."]);
  const json = JSON.stringify(scenario);
  for (let offset = 0, index = 1; offset < json.length; offset += PROJECT_SHEET_CHUNK_SIZE, index += 1) {
    rows.push(["JSON_PART", index, json.slice(offset, offset + PROJECT_SHEET_CHUNK_SIZE)]);
  }
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [
    { wch: 36 }, { wch: 46 }, { wch: 24 }, { wch: 32 },
    { wch: 16 }, { wch: 36 }, { wch: 22 }, { wch: 40 }
  ];
  sheet["!autofilter"] = { ref: `A7:D${7 + Object.keys(scenario.primaryValues).length + Object.keys(scenario.secondaryValues).length}` };
  return sheet;
}

function exportProjectWorkbook() {
  if (!state.workbook || !state.items.length) {
    showNotice("Nejprve načtěte zdrojový Excel, aby bylo možné vyexportovat projekt.", true);
    return;
  }
  try {
    const workbook = state.workbook;
    const existingProjectSheet = workbook.SheetNames.find((name) =>
      workbook.Sheets[name]?.A2?.v === "Formát"
      && workbook.Sheets[name]?.B2?.v === PROJECT_SHEET_MARKER);
    if (existingProjectSheet) {
      delete workbook.Sheets[existingProjectSheet];
      workbook.SheetNames = workbook.SheetNames.filter((name) => name !== existingProjectSheet);
    }
    let projectSheetName = "Projekt aplikace";
    let suffix = 2;
    while (workbook.SheetNames.includes(projectSheetName)) {
      projectSheetName = `Projekt aplikace ${suffix}`;
      suffix += 1;
    }
    const scenario = getProjectScenarioData();
    XLSX.utils.book_append_sheet(workbook, createProjectExportSheet(scenario), projectSheetName);
    const file = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
    const download = document.createElement("a");
    const url = URL.createObjectURL(new Blob([file], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }));
    download.href = url;
    download.download = `${state.fileName.replace(/\.(?:xlsx|xls|xlsm)$/i, "") || "Projekt"}_projekt.xlsx`;
    document.body.append(download);
    download.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      download.remove();
    }, 1000);
    showNotice("Projekt byl vyexportován do Excelu. Při dalším použití importujte tento soubor; obsahuje zdrojový sešit i uložený scénář.");
  } catch (error) {
    showNotice(`Export projektu do Excelu se nepodařil: ${error instanceof Error ? error.message : String(error)}`, true);
  }
}

document.querySelector("#excel-export").addEventListener("click", () => {
  try {
    exportProjectWorkbook();
  } catch (error) {
    showNotice(`Export projektu do Excelu se nepodařil: ${error instanceof Error ? error.message : String(error)}`, true);
  }
});

async function loadFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xls|xlsm)$/i.test(file.name)) {
    showNotice("Vyberte soubor Excel ve formátu .xlsx, .xls nebo .xlsm.", true);
    return;
  }
  clearNotice();
  elements.workbookStatus.textContent = "Načítání…";
  try {
    const loaded = await readWorkbook(file);
    loaded.items.push(...createResidentParkingVariants(loaded.items));
    Object.assign(state, loaded, {
      selectedCategory: "", selectedItemKey: "", search: "", hideUnrealizedItems: false,
      collapsedCategories: new Set(), collapsedLists: {}, projectInputGroupsOpen: {},
      derivedDetailsOpen: false, capacityEdits: {}, overrides: {},
      reviewChecks: {}, standardComparisons: {}, healthChecks: {}, parkingShares: {},
      realizationSelections: {},
      itemEstimateValues: {},
      scenarioChanged: false, inputsChanged: false, persistenceError: false
    });
    const scenarioReadError = readScenario(loaded.projectScenario);
    elements.search.value = state.search;
    elements.excelExport.disabled = false;
    elements.workbookName.textContent = loaded.fileName;
    elements.itemTotal.textContent = String(loaded.items.length);
    renderProjectInputs();
    renderCategories();
    renderRequirements();
    renderIntegrityWarning();
    if (scenarioReadError) showNotice(scenarioReadError, true);
    else showNotice(`Načteno ${loaded.items.length} položek, ${loaded.rules.length} pravidel a ${loaded.inputs.length} definic vstupů. Změny scénáře se ukládají jen v tomto prohlížeči.`);
  } catch (error) {
    elements.workbookStatus.textContent = "Chyba načtení";
    showNotice(error instanceof Error ? error.message : "Sešit se nepodařilo načíst.", true);
  }
}

async function loadDefaultWorkbook() {
  try {
    const response = await fetch(new URL(DEFAULT_WORKBOOK_NAME, document.baseURI));
    if (!response.ok) throw new Error(`Výchozí sešit se nepodařilo načíst (HTTP ${response.status}).`);
    const file = new File([await response.blob()], DEFAULT_WORKBOOK_NAME, { lastModified: 0 });
    await loadFile(file);
  } catch (error) {
    elements.workbookStatus.textContent = "Čeká na načtení";
    showNotice(`${error instanceof Error ? error.message : "Výchozí sešit se nepodařilo načíst."} Vyberte jej ručně tlačítkem „Načíst pracovní Excel“.`, true);
  }
}

function resetScenario() {
  state.primaryValues = { ...state.defaults.primary };
  state.secondaryValues = { ...state.defaults.secondary };
  state.capacityEdits = {};
  state.overrides = {};
  state.reviewChecks = {};
  state.standardComparisons = {};
  state.healthChecks = {};
  state.parkingShares = {};
  state.realizationSelections = {};
  state.itemEstimateValues = {};
  state.hideUnrealizedItems = false;
  state.collapsedCategories = new Set();
  state.collapsedLists = {};
  state.projectInputGroupsOpen = {};
  state.derivedDetailsOpen = false;
  state.selectedCategory = "";
  state.selectedItemKey = "";
  state.search = "";
  elements.search.value = "";
  state.scenarioChanged = false;
  state.inputsChanged = false;
  state.persistenceError = false;
  try {
    if (state.fileKey) localStorage.removeItem(`territory-scenario:${state.fileKey}`);
  } catch (error) {
    showNotice(`Uložené změny se nepodařilo odstranit: ${error.message}`, true);
    return;
  }
  renderProjectInputs();
  renderCategories();
  renderRequirements();
  showNotice("Hodnoty byly obnoveny podle právě načteného sešitu.");
}

elements.workbookInput.addEventListener("change", (event) => {
  loadFile(event.currentTarget.files?.[0]);
  event.currentTarget.value = "";
});
elements.dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  elements.dropzone.classList.add("is-dragging");
});
elements.dropzone.addEventListener("dragleave", () => elements.dropzone.classList.remove("is-dragging"));
elements.dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  elements.dropzone.classList.remove("is-dragging");
  loadFile(event.dataTransfer?.files?.[0]);
});
elements.projectInputGrid.addEventListener("change", handleInputChange);
elements.secondaryInputGrid.addEventListener("change", handleInputChange);
elements.legalList.addEventListener("change", handleInlineProjectInputChange);
elements.recommendedList.addEventListener("change", handleInlineProjectInputChange);
elements.infrastructureList.addEventListener("change", handleInlineProjectInputChange);
elements.reviewList.addEventListener("change", handleInlineProjectInputChange);
elements.categoryList.addEventListener("click", (event) => {
  const collapseButton = event.target.closest("button[data-collapse-category]");
  if (collapseButton) {
    const category = collapseButton.dataset.collapseCategory;
    if (state.collapsedCategories.has(category)) state.collapsedCategories.delete(category);
    else state.collapsedCategories.add(category);
    renderCategories();
    saveScenario(false);
    return;
  }
  const equipmentButton = event.target.closest("button[data-item-key]");
  if (equipmentButton) {
    const clickedItemKey = equipmentButton.dataset.itemKey;
    state.selectedItemKey = state.selectedItemKey === clickedItemKey ? "" : clickedItemKey;
    state.selectedCategory = state.items.find((item) => item.key === state.selectedItemKey)?.category || "";
    state.search = "";
    elements.search.value = "";
    renderCategories();
    renderRequirements();
    saveScenario(false);
    return;
  }
  const button = event.target.closest("button[data-category]");
  if (!button) return;
  const clickedCategory = button.dataset.category;
  state.selectedCategory = state.selectedCategory === clickedCategory ? "" : clickedCategory;
  state.selectedItemKey = "";
  state.search = "";
  elements.search.value = "";
  renderCategories();
  renderRequirements();
  saveScenario(false);
});
elements.requirementColumns.addEventListener("click", (event) => {
  const realizationButton = event.target.closest("button[data-realization-item-key]");
  if (realizationButton) {
    event.preventDefault();
    event.stopPropagation();
    const item = state.items.find((candidate) =>
      candidate.key === realizationButton.dataset.realizationItemKey);
    if (item) setRealizationSelection(item, realizationButton.getAttribute("aria-pressed") !== "true");
    return;
  }
  const button = event.target.closest("button[data-collapse-list]");
  if (!button) return;
  const list = document.getElementById(button.getAttribute("aria-controls"));
  if (!list) return;
  list.hidden = !list.hidden;
  const expanded = !list.hidden;
  const listName = button.dataset.collapseList;
  state.collapsedLists[listName] = !expanded;
  const sectionName = button.closest(".requirement-column").querySelector("h3").textContent;
  button.setAttribute("aria-expanded", String(expanded));
  button.setAttribute("aria-label", `${expanded ? "Sbalit" : "Rozbalit"} ${sectionName}`);
  saveScenario(false);
});
elements.search.addEventListener("input", () => {
  state.search = elements.search.value;
  state.selectedItemKey = "";
  renderRequirements();
  saveScenario(false);
});
elements.hideUnrealizedItems.addEventListener("change", () => {
  state.hideUnrealizedItems = elements.hideUnrealizedItems.checked;
  const matching = matchingItems();
  const selectedKeys = getRealizationPlan().selectedKeys;
  const count = state.hideUnrealizedItems
    ? matching.filter((item) => selectedKeys.has(item.key)).length : matching.length;
  elements.visibleCount.textContent = state.hideUnrealizedItems
    ? `${count} realizovaných` : `${count} položek`;
  updateRequirementVisibility();
  saveScenario(false);
});
document.querySelector(".derived-details").addEventListener("toggle", (event) => {
  state.derivedDetailsOpen = event.currentTarget.open;
  saveScenario(false);
});
elements.requirementColumns.addEventListener("input", updateTargetedPlanInput);
elements.categoryList.addEventListener("change", updateRealizationSelection);
elements.legalList.addEventListener("input", updateCapacityEdit);
elements.legalList.addEventListener("change", updateCapacityEdit);
elements.legalList.addEventListener("input", updateHealthcareCheck);
elements.legalList.addEventListener("change", updateHealthcareCheck);
elements.legalList.addEventListener("input", updateManualReview);
elements.legalList.addEventListener("change", updateManualReview);
elements.legalList.addEventListener("input", updateStandardComparison);
elements.legalList.addEventListener("change", updateStandardComparison);
elements.legalList.addEventListener("click", handleRequirementClick);
elements.recommendedList.addEventListener("input", updateCapacityEdit);
elements.recommendedList.addEventListener("change", updateCapacityEdit);
elements.recommendedList.addEventListener("input", updateHealthcareCheck);
elements.recommendedList.addEventListener("change", updateHealthcareCheck);
elements.recommendedList.addEventListener("input", updateManualReview);
elements.recommendedList.addEventListener("change", updateManualReview);
elements.recommendedList.addEventListener("input", updateStandardComparison);
elements.recommendedList.addEventListener("change", updateStandardComparison);
elements.recommendedList.addEventListener("click", handleRequirementClick);
elements.infrastructureList.addEventListener("input", updateCapacityEdit);
elements.infrastructureList.addEventListener("change", updateCapacityEdit);
elements.infrastructureList.addEventListener("input", updateManualReview);
elements.infrastructureList.addEventListener("change", updateManualReview);
elements.infrastructureList.addEventListener("change", handleInlineProjectInputChange);
elements.infrastructureList.addEventListener("click", handleRequirementClick);
elements.reviewList.addEventListener("input", updateCapacityEdit);
elements.reviewList.addEventListener("change", updateCapacityEdit);
elements.reviewList.addEventListener("input", updateManualReview);
elements.reviewList.addEventListener("change", updateManualReview);
elements.reviewList.addEventListener("input", updateHealthcareCheck);
elements.reviewList.addEventListener("change", updateHealthcareCheck);
elements.reviewList.addEventListener("change", handleInlineProjectInputChange);
elements.reviewList.addEventListener("click", handleRequirementClick);
elements.resetInputs.addEventListener("click", resetScenario);
document.querySelector("#open-feedback-report").addEventListener("click", () => {
  const report = document.querySelector("#feedback-report");
  report.open = true;
  report.scrollIntoView({ behavior: "auto", block: "start" });
  document.querySelector("#feedback-errors").focus({ preventScroll: true });
});
document.querySelector("#feedback-form").addEventListener("submit", exportFeedbackReport);
function applyTheme(theme) {
  const isNight = theme === "night";
  document.documentElement.dataset.theme = isNight ? "night" : "day";
  elements.themeToggle.querySelector("span").textContent = isNight ? "Denní režim" : "Noční režim";
  elements.themeToggle.setAttribute("aria-pressed", String(isNight));
  document.querySelector('meta[name="theme-color"]').content = isNight ? "#080808" : "#f5f6f1";
}

elements.themeToggle.addEventListener("click", () => {
  const theme = document.documentElement.dataset.theme === "night" ? "day" : "night";
  applyTheme(theme);
  try {
    localStorage.setItem("territory-theme", theme);
  } catch (error) {
    showNotice(`Nastavení motivu se nepodařilo uložit: ${error.message}`, true);
  }
});
try {
  const savedTheme = localStorage.getItem("territory-theme");
  if (savedTheme === "day" || savedTheme === "night") applyTheme(savedTheme);
} catch (error) {
  showNotice(`Uložené nastavení motivu se nepodařilo načíst: ${error.message}`, true);
}
document.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
    event.preventDefault();
    elements.search.focus();
  }
});

loadDefaultWorkbook();
