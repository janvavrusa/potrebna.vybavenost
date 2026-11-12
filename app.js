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
  { row: 14, code: "REZIM_STD", label: "Zvolená hladina navrhovaného standardu", kind: "select",
    options: ["ZAKON", "STANDARD", "MAX_STANDARD"],
    labels: ["Zákonné minimum", "Doporučený standard", "Nejvyšší standard"], unit: "" },
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
  { row: 19, code: "O_prah", label: "Současný počet obyvatel obce pro standardy sociálních služeb", unit: "obyvatel", formula: "O_sidlo (současní obyvatelé) + O_celk (přírůstek) + O_spad" },
  { row: null, code: "BJ_pocet", label: "Počet nových bytových jednotek", unit: "jednotek", formula: "Ruční počet; jinak celková HPP / vážená plocha jednotky; jinak O_celk / počet osob na byt" },
  { row: null, code: "HPP_byty", label: "Celková HPP navrženého bydlení", unit: "m² HPP", formula: "BJ_pocet × HPP na byt" },
  { row: 20, code: "rho_hruba", label: "Hrubá obytná hustota", unit: "obyv./ha", formula: "O_celk / S_uzemi" },
  { row: 21, code: "rho_cista", label: "Čistá obytná hustota", unit: "obyv./ha", formula: "O_celk / S_zastav" },
  { row: 22, code: "k_tvar", label: "Tvarový koeficient", unit: "", formula: "Konstanty_a_Ciselniky: tvarový koeficient" },
  { row: 23, code: "k_pesi", label: "Uliční koeficient docházky", unit: "", formula: "Konstanty_a_Ciselniky: uliční koeficient" },
  { row: 24, code: "c_tvar", label: "Korekční činitel sítě bodů", unit: "", formula: "Konstanty_a_Ciselniky: korekce sítě bodů" },
  { row: 25, code: "R_max", label: "Maximální poloměr sítě", unit: "m", formula: "k_tvar × √(S_uzemi × 10 000)" },
  { row: 26, code: "L_max", label: "Max. pěší vzdálenost z těžiště", unit: "m", formula: "k_pesi × √(S_uzemi × 10 000)" }
];
const DERIVED_PARAMETER_INPUTS = [
  { row: null, code: "BJ_POCET_MANUAL", label: "Počet nových bytových jednotek (byty i rodinné domy)", kind: "number", min: 0, step: 1,
    defaultValue: "", unit: "bytových jednotek", note: "Volitelný přepis automatického odhadu; ponechte prázdné pro odhad z obyvatel nebo z celkové HPP a průměrné plochy." },
  { row: null, code: "PARKING_TOTAL_HPP", label: "HPP pro odhad počtu bytů (volitelné)", kind: "number", min: 0, step: 1,
    defaultValue: "", unit: "m² HPP", note: "Volitelný vstup pro odhad počtu bytů. Výsledná HPP navrženého bydlení se zobrazuje jako odvozený ukazatel." }
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
    row: null, code: "POLY_DISCOUNT_COEFFICIENT", label: "Koeficient efektivity společné polikliniky",
    kind: "number", min: 0, max: 1, step: 0.05, defaultValue: 0.7,
    unit: "násobek součtu ploch a investic",
    note: "Výchozí scénář 0,7 snižuje součet modelových ploch a investic díky sdíleným prostorům. Koeficient je editovatelný; kapacita ordinací se nesnižuje."
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
  { code: "SOCIAL_DEMOGRAPHIC_COEFFICIENT", label: "Demografický koeficient kapacity", kind: "number", min: 0, max: 1, step: 0.00001, color: "area" },
  { code: "ITEM_USABLE_AREA_PER_UNIT", label: "Užitná plocha na jednotku", kind: "number", min: 0, step: 0.1, unit: "m² / jednotku", color: "area" },
  { code: "ITEM_TOTAL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, unit: "m² celkem / m² užitná", color: "floor" },
  { code: "ITEM_VOLUME_PER_AREA", label: "Koeficient obestavěného prostoru", kind: "number", min: 0, step: 0.1, unit: "m³ / m² celkem", color: "volume" },
  { code: "ITEM_SITE_AREA_PER_UNIT", label: "Plocha pozemku na jednotku", kind: "number", min: 0, step: 0.1, unit: "m² / jednotku", color: "site" },
  { code: "ITEM_PAVED_AREA_PER_UNIT", label: "Zpevněná plocha na jednotku", kind: "number", min: 0, step: 0.1, unit: "m² / jednotku", color: "paved" },
  { code: "PUBLIC_SPACE_QUALIFIED_AREA_HA", label: "Kvalifikovaná zastavitelná plocha", kind: "number", min: 0, step: 0.1, unit: "ha", color: "area", note: "Výchozí hodnota přebírá S_zastav; upravte ji na související plochy bydlení, rekreace, občanského vybavení a smíšené obytné bez pozemních komunikací." },
  { code: "PUBLIC_SPACE_PAVED_SHARE", label: "Podíl plochy veřejného prostranství oceňovaný jako zpevněný", kind: "percent", min: 0, max: 1, step: 0.05, unit: "%", color: "paved", note: "Modelový předpoklad skladby plochy, nikoli požadavek vyhlášky; upravte podle návrhu." },
  { code: "PUBLIC_SPACE_PAVING_RATE_M2", label: "Modelová cena zpevněné plochy", kind: "number", min: 0, step: 100, unit: "Kč / m²", color: "rts", note: "Orientační scénářová sazba mimo RTS; ověřte skladbu konstrukce, povrch a místní rozpočet." },
  { code: "PARK_GREEN_AREA_HA", label: "Modelová plocha jednoho parku", kind: "number", min: 0.5, step: 0.1, unit: "ha / park", color: "site", note: "MMR doporučuje parkovou zeleň v rozsahu 0,5–1 ha; výchozí minimum 0,5 ha je editovatelné." },
  { code: "PARK_LANDSCAPING_RATE_M2", label: "Agregovaná cena parkové úpravy", kind: "number", min: 0, step: 100, unit: "Kč / m²", color: "rts", note: "Editovatelný souhrnný scénář mimo RTS zahrnující parkovou úpravu; upravte podle místních sazeb a rozsahu." },
  { code: "ITEM_RTS_RATE", label: "Cena dle měrné jednotky", kind: "number", min: 0, step: 100, unit: "Kč / měrná jednotka", color: "rts" },
  { code: "ITEM_RTS_RATE_BASIS", label: "Základ ceny", kind: "select", options: ["volume", "unit"], labels: ["Kč / m³ obestavěného prostoru", "Kč / měrnou jednotku"], unit: "", color: "rts" },
  { code: "ITEM_RTS_FIXED_COST", label: "Další fixní náklady dle modelu", kind: "number", min: 0, step: 10000, unit: "Kč / projekt", color: "fixed" },
  { code: "ITEM_COST_SURCHARGE", label: "VRN a rezerva", kind: "percent", min: 0, max: 2, step: 0.01, unit: "%; z nákladů ZRN", color: "surcharge" }
];
const GENERIC_ITEM_ESTIMATE_CODES = new Set([
  "ITEM_QUANTITY", "ITEM_USABLE_AREA_PER_UNIT", "ITEM_TOTAL_FLOOR_FACTOR",
  "ITEM_VOLUME_PER_AREA", "ITEM_SITE_AREA_PER_UNIT", "ITEM_PAVED_AREA_PER_UNIT",
  "ITEM_RTS_RATE", "ITEM_RTS_RATE_BASIS", "ITEM_RTS_FIXED_COST", "ITEM_COST_SURCHARGE"
]);
const SUPERMARKET_ESTIMATE_INPUTS = [
  { code: "SUPERMARKET_MIN_POPULATION", label: "Minimální počet obyvatel pro scénář supermarketu", kind: "number", min: 0, step: 100, unit: "obyvatel", color: "area", note: "Výchozí pracovní práh 3 000 obyvatel; nejde o pravidlo MMR ani zákonnou povinnost." },
  { code: "SUPERMARKET_POPULATION_PER_STORE", label: "Přírůstek obyvatel na jednu modelovou prodejnu", kind: "number", min: 1, step: 100, unit: "obyvatel/prodejnu", color: "area" },
  ...ITEM_ESTIMATE_INPUTS.filter((field) => [
    "ITEM_USABLE_AREA_PER_UNIT", "ITEM_TOTAL_FLOOR_FACTOR", "ITEM_VOLUME_PER_AREA",
    "ITEM_SITE_AREA_PER_UNIT", "ITEM_PAVED_AREA_PER_UNIT", "ITEM_RTS_RATE",
    "ITEM_RTS_RATE_BASIS", "ITEM_RTS_FIXED_COST", "ITEM_COST_SURCHARGE"
  ].includes(field.code))
];
const PARKING_ESTIMATE_INPUTS = [
  { code: "PARKING_BUILDING_SHARE", label: "Podíl stání v parkovacím domě (zbytek na zpevněné ploše)", kind: "percent", min: 0, max: 1, step: 0.05, unit: "%", color: "site", note: "Zadejte procento stání umístěných v parkovacím domě; zbývající podíl do 100 % se počítá jako venkovní stání na zpevněné ploše. Výchozí hodnoty jsou 20 % zákonné a 80 % doporučené varianty." },
  { code: "PARKING_VOLUME_PER_STALL_M3", label: "Obestavěný prostor parkovacího domu na stání", kind: "number", min: 0, step: 1, unit: "m³ / stání", color: "volume", note: "Pracovní předpoklad převzatý z nákladového modelu sešitu; ověřte jej podle návrhu objektu." },
  { code: "PARKING_BUILDING_RATE_M3", label: "Sazba parkovacího domu", kind: "number", min: 0, step: 100, unit: "Kč / m³ OP", color: "rts", note: "Výchozí sazba je z RTS 2026, JKSO 811.5; upravte konstrukční charakteristiku a sazbu dle projektu." },
  { code: "PARKING_SURFACE_RATE_M3", label: "Modelová sazba venkovního parkování", kind: "number", min: 0, step: 100, unit: "Kč / m³ ekv.", color: "rts", note: "Pracovní sazba 7 500 Kč/m³ z nákladového modelu sešitu; jde o objemový ekvivalent, nikoli položkovou cenu zpevněné plochy." },
  { code: "PARKING_SURFACE_AREA_PER_STALL_M2", label: "Zpevněná plocha na venkovní stání", kind: "number", min: 0, step: 1, unit: "m² / stání", color: "paved" },
  ITEM_ESTIMATE_INPUTS.find((field) => field.code === "ITEM_COST_SURCHARGE")
];
const ALTERNATIVE_ESTIMATE_INPUTS = [
  { code: "ALT_RAINFALL_DEPTH_MM", label: "Modelová srážka", kind: "number", min: 0, step: 1, unit: "mm / srážková událost", color: "area" },
  { code: "ALT_RUNOFF_COEFFICIENT", label: "Součinitel odtoku", kind: "number", min: 0, max: 1, step: 0.05, unit: "bez jednotky", color: "paved" },
  { code: "ALT_ROOF_RUNOFF_COEFFICIENT", label: "Součinitel odtoku ze střech", kind: "number", min: 0, max: 1, step: 0.05, unit: "bez jednotky", color: "paved" },
  { code: "ALT_PAVED_RUNOFF_COEFFICIENT", label: "Součinitel odtoku ze zpevněných ploch", kind: "number", min: 0, max: 1, step: 0.05, unit: "bez jednotky", color: "paved" },
  { code: "ALT_RAINWATER_RATE", label: "Odhad ceny retenční nádrže včetně stavebních prací", kind: "number", min: 0, step: 100, unit: "Kč / m³ retenčního objemu", color: "rts" },
  { code: "ALT_WATER_LITERS_PER_PERSON_DAY", label: "Spotřeba vody na obyvatele", kind: "number", min: 0, step: 1, unit: "l / osoba / den", color: "area" },
  { code: "ALT_FACILITY_WATER_LITERS_PER_M3_DAY", label: "Spotřeba vody realizované vybavenosti", kind: "number", min: 0, step: 0.1, unit: "l / m³ obestavěného prostoru / den", color: "volume", note: "Výchozí objemový odhad 1,5 l/m³/den je společný plánovací proxy, nikoli norma; upravte podle druhu provozu a konkrétního projektu." },
  { code: "ALT_PEAK_DAY_FACTOR", label: "Součinitel špičkového dne", kind: "number", min: 1, step: 0.1, unit: "násobek průměrného dne", color: "volume" },
  { code: "ALT_WASTEWATER_RETURN", label: "Podíl vody odváděný jako splašky", kind: "percent", min: 0, max: 1, step: 0.01, unit: "% z dodané vody", color: "paved" },
  { code: "ALT_NETWORK_LENGTH_M", label: "Délka nové trasy potrubí nebo kabelu", kind: "number", min: 0, step: 50, unit: "m", color: "site" },
  { code: "ALT_WATER_PIPE_RATE", label: "Základní cena vodovodního potrubí včetně pokládky", kind: "number", min: 0, step: 100, unit: "Kč / m", color: "rts", note: "Základní sazba při rezidenční návrhové kapacitě; cena trasy se škáluje odmocninou poměru celkové a rezidenční kapacity jako hrubý odhad větší dimenze." },
  { code: "ALT_WATER_METER_CHAMBER_COUNT", label: "Počet vodoměrných a armaturních šachet", kind: "number", min: 0, step: 1, unit: "ks", color: "fixed" },
  { code: "ALT_WATER_METER_CHAMBER_RATE", label: "Odhad ceny jedné vodoměrné a armaturní šachty", kind: "number", min: 0, step: 1000, unit: "Kč / ks", color: "rts" },
  { code: "ALT_SEWER_PIPE_RATE", label: "Základní cena splaškového potrubí včetně pokládky", kind: "number", min: 0, step: 100, unit: "Kč / m", color: "rts", note: "Základní sazba při rezidenční návrhové kapacitě; cena trasy se škáluje odmocninou poměru celkové a rezidenční kapacity jako hrubý odhad větší dimenze." },
  { code: "ALT_SEWER_MANHOLE_SPACING_M", label: "Modelová rozteč revizních šachet", kind: "number", min: 1, step: 5, unit: "m", color: "site" },
  { code: "ALT_SEWER_MANHOLE_RATE", label: "Odhad ceny jedné revizní šachty", kind: "number", min: 0, step: 1000, unit: "Kč / ks", color: "rts" },
  { code: "ALT_ENERGY_DENSITY_W_M2", label: "Modelová elektrická zátěž budov", kind: "number", min: 0, step: 5, unit: "W / m² celkové plochy", color: "volume" },
  { code: "ALT_ENERGY_DENSITY_W_M3", label: "Modelový elektrický příkon na objem budov", kind: "number", min: 0, step: 1, unit: "W / m³ obestavěného prostoru", color: "volume", note: "Výchozích 12 W/m³ přibližně odpovídá dřívějšímu odhadu 50 W/m² při 4,2 m³/m²; jde o editovatelný proxy pro objekty bez energetických dat." },
  { code: "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO", label: "Poměr špičkového a průměrného elektrického příkonu", kind: "number", min: 1, step: 0.1, unit: "špička / průměr", color: "volume" },
  { code: "BOILER_HEAT_DENSITY_KWH_M2_Y", label: "Náhradní roční potřeba tepla pro objekty bez energetického údaje", kind: "number", min: 0, step: 5, unit: "kWh / m² / rok", color: "volume" },
  { code: "BOILER_HEAT_DENSITY_KWH_M3_Y", label: "Náhradní roční potřeba tepla na objem budovy", kind: "number", min: 0, step: 1, unit: "kWh / m³ obestavěného prostoru / rok", color: "volume", note: "Výchozích 14,3 kWh/m³/rok přibližně odpovídá 60 kWh/m²/rok při 4,2 m³/m²; jde o editovatelný proxy, ne normovou hodnotu." },
  { code: "BOILER_HEATING_SHARE", label: "Podíl neelektrické energie uvažovaný pro ústřední vytápění", kind: "percent", min: 0, max: 1, step: 0.05, unit: "% neelektrické energie", color: "area" },
  { code: "BOILER_FULL_LOAD_HOURS", label: "Ekvivalentní roční hodiny plného zatížení kotelny", kind: "number", min: 1, step: 100, unit: "h / rok", color: "volume" },
  { code: "BOILER_DESIGN_RESERVE", label: "Součinitel návrhové rezervy výkonu", kind: "number", min: 1, step: 0.05, unit: "násobek", color: "volume" },
  { code: "BOILER_UNIT_CAPACITY_KW", label: "Kapacita jednoho kotlového modulu", kind: "number", min: 1, step: 50, unit: "kW tepelného výkonu / modul", color: "volume" },
  { code: "BOILER_UNIT_COST", label: "Modelový odhad ceny jednoho kotlového modulu", kind: "number", min: 0, step: 10000, unit: "Kč / modul; odhad 10 000 Kč/kW", color: "rts" },
  { code: "ALT_POWER_CABLE_RATE", label: "Základní cena elektrického kabelu včetně pokládky", kind: "number", min: 0, step: 100, unit: "Kč / m", color: "rts", note: "Základní sazba při špičkovém příkonu základního bydlení; odhad ceny kabelu se mění přímo úměrně poměru celkového a rezidenčního příkonu." },
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
const BRNO_CAR_USE_RATES = [
  { key: "housing", label: "Bydlení", rate: 120, bound: 90, visitor: 10, dwellings: true },
  { key: "retail-ground", label: "Jednotlivé obchody v parteru", rate: 70, bound: 10, visitor: 90 },
  { key: "small-service", label: "Služby a drobné provozovny", rate: 40, bound: 10, visitor: 90 },
  { key: "auto-service", label: "Autoservis", rate: 20, bound: 10, visitor: 90 },
  { key: "large-retail", label: "Velkoplošný obchod a služby", rate: 40, bound: 10, visitor: 90 },
  { key: "admin-low", label: "Administrativa s malou návštěvností", rate: 50, bound: 90, visitor: 10 },
  { key: "admin-high", label: "Administrativa s velkou návštěvností", rate: 45, bound: 60, visitor: 40 },
  { key: "long-stay", label: "Dlouhodobé ubytování", rate: 120, bound: 80, visitor: 20 },
  { key: "student-housing", label: "Ubytování pro studenty", rate: 250, bound: 90, visitor: 10 },
  { key: "short-stay", label: "Krátkodobé ubytování", rate: 100, bound: 90, visitor: 10 },
  { key: "motel", label: "Motel", rate: 25, bound: 90, visitor: 10 },
  { key: "hostel", label: "Hostel", rate: 180, bound: 90, visitor: 10 },
  { key: "social-care", label: "Institucionální a sociální péče", rate: 350, bound: 35, visitor: 65 },
  { key: "school", label: "Školství", rate: 250, bound: 30, visitor: 70 },
  { key: "nursery", label: "Jesle a mateřská škola", rate: 300, bound: 30, visitor: 70 },
  { key: "university", label: "Vysoká škola", rate: 100, bound: 80, visitor: 20 },
  { key: "training", label: "Vzdělávání a kongres", rate: 60, bound: 10, visitor: 90 },
  { key: "gathering", label: "Shromažďovací prostory", rate: 60, bound: 20, visitor: 80 },
  { key: "church", label: "Kostel a modlitebna", rate: 200, bound: 5, visitor: 95 },
  { key: "ceremony", label: "Obřadní síň a krematorium", rate: 120, bound: 10, visitor: 90 },
  { key: "culture", label: "Kulturní instituce", rate: 120, bound: 20, visitor: 80 },
  { key: "outpatient-health", label: "Ambulantní zdravotnické zařízení", rate: 50, bound: 30, visitor: 70 },
  { key: "inpatient-health", label: "Lůžkové zdravotnické zařízení", rate: 300, bound: 50, visitor: 50 },
  { key: "indoor-sport", label: "Sportovní centrum s vnitřní hrací plochou", rate: 100, bound: 20, visitor: 80 },
  { key: "sport-service", label: "Sportovní centrum bez hrací plochy a bazény", rate: 40, bound: 10, visitor: 90 },
  { key: "outdoor-sport", label: "Venkovní sportoviště bez diváků", rate: 120, bound: 10, visitor: 90 },
  { key: "football", label: "Fotbalové hřiště", rate: 400, bound: 10, visitor: 90 },
  { key: "production", label: "Výroba (zvolit ukazatel 200–800 m²/stání)", rate: null, bound: 10, visitor: 90 },
  { key: "storage", label: "Skladování", rate: 200, bound: 30, visitor: 70 }
];
const BRNO_BIKE_USE_RATES = [
  { key: "housing", label: "Bydlení", rate: 60, bound: 90, visitor: 10 },
  { key: "retail-ground", label: "Jednotlivé obchody v parteru", rate: 150, bound: 40, visitor: 60 },
  { key: "small-service", label: "Služby a drobné provozovny", rate: 150, bound: 30, visitor: 70 },
  { key: "auto-service", label: "Autoservis", rate: 150, bound: 30, visitor: 70 },
  { key: "large-retail", label: "Velkoplošný obchod a služby", rate: 250, bound: 30, visitor: 70 },
  { key: "admin-low", label: "Administrativa s malou návštěvností", rate: 200, bound: 90, visitor: 10 },
  { key: "admin-high", label: "Administrativa s velkou návštěvností", rate: 150, bound: 60, visitor: 40 },
  { key: "long-stay", label: "Dlouhodobé ubytování", rate: 60, bound: 90, visitor: 10 },
  { key: "student-housing", label: "Ubytování pro studenty", rate: 60, bound: 90, visitor: 10 },
  { key: "short-stay", label: "Krátkodobé ubytování", rate: 1000, bound: 90, visitor: 10 },
  { key: "motel", label: "Motel", rate: 1000, bound: 90, visitor: 10 },
  { key: "hostel", label: "Hostel", rate: 1000, bound: 90, visitor: 10 },
  { key: "social-care", label: "Institucionální a sociální péče", rate: 500, bound: 30, visitor: 70 },
  { key: "youth-care", label: "Domovy mládeže a azylové domy", rate: 60, bound: 30, visitor: 70 },
  { key: "school", label: "Školství", rate: 60, bound: 90, visitor: 10 },
  { key: "nursery", label: "Jesle a mateřská škola", rate: 150, bound: 40, visitor: 60 },
  { key: "university", label: "Vysoká škola", rate: 60, bound: 70, visitor: 30 },
  { key: "training", label: "Vzdělávání a kongres", rate: 300, bound: 80, visitor: 20 },
  { key: "gathering", label: "Shromažďovací prostory", rate: 60, bound: 30, visitor: 70 },
  { key: "church", label: "Kostel a modlitebna", rate: 60, bound: 30, visitor: 70 },
  { key: "ceremony", label: "Obřadní síň a krematorium", rate: 60, bound: 30, visitor: 70 },
  { key: "culture", label: "Kulturní instituce", rate: 250, bound: 30, visitor: 70 },
  { key: "outpatient-health", label: "Ambulantní zdravotnické zařízení", rate: 250, bound: 70, visitor: 30 },
  { key: "inpatient-health", label: "Lůžkové zdravotnické zařízení", rate: 500, bound: 70, visitor: 30 },
  { key: "indoor-sport", label: "Sportovní centrum s vnitřní hrací plochou", rate: 100, bound: 40, visitor: 60 },
  { key: "sport-service", label: "Sportovní centrum bez hrací plochy a bazény", rate: 40, bound: 40, visitor: 60 },
  { key: "outdoor-sport", label: "Venkovní sportoviště bez diváků", rate: 120, bound: 10, visitor: 90 },
  { key: "football", label: "Fotbalové hřiště", rate: 400, bound: 10, visitor: 90 },
  { key: "production", label: "Výroba", rate: 500, bound: 90, visitor: 10 },
  { key: "storage", label: "Skladování", rate: 500, bound: 90, visitor: 10 }
];
const BRNO_HPP_ITEM_IDS = {
  "retail-ground": ["VYB_13A"],
  "small-service": ["VYB_12"],
  "large-retail": ["VYB_13B"],
  "admin-low": ["VYB_11"],
  "admin-high": [],
  "long-stay": [],
  "social-care": ["VYB_06A", "VYB_06B", "VYB_06C", "VYB_07A_1", "VYB_07A_2"],
  "school": ["VYB_02A", "VYB_02B"],
  "nursery": ["VYB_01", "VYB_01_J"],
  "training": ["VYB_02C_1", "VYB_02C_2"],
  "gathering": ["VYB_10B_1"],
  "culture": ["VYB_10A", "VYB_10B_2"],
  "outpatient-health": [
    "VYB_09A_4", "VYB_09B_1", "VYB_09B_2", "VYB_16", "VYB_17",
    "VYB_18", "VYB_19", "VYB_20"
  ],
  "indoor-sport": ["VYB_04", "VYB_05B"]
};
const BRNO_REQUIREMENTS = [
  {
    key: "BSP_BRNO_CAR_PARKING",
    label: "Parkování motorových vozidel",
    type: "cars",
    description: "Vypočítá minimum stání z HPP realizací a zóny podle mapy v příloze č. 3."
  },
  {
    key: "BSP_BRNO_BIKE_PARKING",
    label: "Parkování jízdních kol",
    type: "bikes",
    description: "Počítá vázaná a návštěvnická místa z HPP právě vybraných realizací."
  },
  {
    key: "BSP_BRNO_TREE_SPACES",
    label: "Prostory pro růst stromů u povrchového parkování",
    type: "trees",
    description: "Navrhuje jeden prostor pro růst stromu na každých osm venkovních stání."
  }
];
const BRNO_REQUIREMENT_KEYS = new Set(BRNO_REQUIREMENTS.map(({ key }) => key));
const REMOVED_BRNO_REQUIREMENT_KEYS = new Set([
  "BSP_BRNO_BOUNDARY_SETBACK", "BSP_BRNO_STORMWATER_LIMIT"
]);
const USER_ADDED_ITEM_IDS = new Set([
  "VYB_05B", "VYB_11", "VYB_13A", "VYB_13B", HEATING_BOILER_ID
]);
const VOLUNTEER_FIRE_STATION_ID = "VYB_32";
const RESIDENT_PARKING_ID = "VYB_15A_1";
const RESIDENT_PARKING_LAW_ID = "VYB_15A_1_LAW";
const RESIDENT_PARKING_RECOMMENDED_ID = "VYB_15A_1_RECOMMENDED";
const RESIDENT_PARKING_RESERVE_ID = "VYB_15A_1_RESERVE";
const SPECIAL_BUILDING_PARKING_KEY = "SPECIAL_BUILDING_PARKING";
const SPECIAL_BUILDING_PARKING_REALIZATION_KEY = "__SPECIAL_BUILDING_PARKING_REALIZATION__";
const SPECIAL_BUILDING_PARKING_RATES = [
  { code: "RETAIL", label: "Obchody a supermarket", m2PerStall: 30, salesAreaShare: 0.8, itemIds: ["VYB_13A", "VYB_13B"] },
  { code: "EDUCATION", label: "Školy a školská zařízení", m2PerStall: 50, itemIds: ["VYB_01", "VYB_02A", "VYB_02B", "VYB_02C_1", "VYB_02C_2", "VYB_04"] },
  { code: "HEALTH", label: "Zdravotnická zařízení", m2PerStall: 50, itemIds: ["VYB_09A_4", "VYB_09B_1", "VYB_09B_2", "VYB_16", "VYB_17", "VYB_18", "VYB_19"] },
  { code: "SOCIAL", label: "Sociální služby", m2PerStall: 100, itemIds: ["VYB_06A", "VYB_06B", "VYB_07A_1", "VYB_07A_2"] },
  { code: "CULTURE", label: "Kultura, krytý sport a veřejné služby", m2PerStall: 50, itemIds: ["VYB_04", "VYB_05B", "VYB_10A", "VYB_10B_1", "VYB_10B_2", "VYB_11", "VYB_12", "VYB_09D", "VYB_16A", VOLUNTEER_FIRE_STATION_ID, "VYB_17A"] }
];
const RESIDENT_PARKING_IDS = new Set([
  RESIDENT_PARKING_ID, RESIDENT_PARKING_LAW_ID,
  RESIDENT_PARKING_RECOMMENDED_ID, RESIDENT_PARKING_RESERVE_ID
]);
const PARKING_SPLIT_ITEM_IDS = new Set([
  ...RESIDENT_PARKING_IDS, "VYB_15A_2", "VYB_15A_3"
]);
const TECHNICAL_ALTERNATIVE_IDS = new Set(["VYB_20A", "VYB_20B_1", "VYB_20B_2", "VYB_20B_3"]);
const INFRASTRUCTURE_ITEM_IDS = TECHNICAL_ALTERNATIVE_IDS;
const INFRASTRUCTURE_RESERVE_SUFFIX = "_RESERVE";
const TECHNICAL_INFRASTRUCTURE_CATEGORY = "TECHNICKÁ INFRASTRUKTURA A VODNÍ HOSPODÁŘSTVÍ";
const RECOMMENDED_SCENARIO_ITEM_IDS = new Set([
  "VYB_11", "VYB_05B", "VYB_02C_1", "VYB_02C_2",
  "VYB_16A", VOLUNTEER_FIRE_STATION_ID, "VYB_17A"
]);
const REALIZATION_NETWORK_IDS = new Set(["VYB_20B_1", "VYB_20B_2", "VYB_20B_3"]);
const REALIZATION_PARKING_IDS = new Set([
  "VYB_15A_1", RESIDENT_PARKING_LAW_ID, RESIDENT_PARKING_RECOMMENDED_ID,
  RESIDENT_PARKING_RESERVE_ID,
  "VYB_15A_2", "VYB_15A_3", "VYB_15B_1", "VYB_15B_2"
]);
const REALIZATION_NON_BUILDING_IDS = new Set(["VYB_14", "VYB_07B"]);
const REALIZATION_OUTDOOR_IDS = new Set(["VYB_03A", "VYB_03B", "VYB_03C", "VYB_04B"]);
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
    name: "Maloobchod",
    context: "Odděluje průběžný místní nákup od většího týdenního nákupu v supermarketu."
  },
  VYB_13B: {
    name: "Supermarket",
    context: "Uživatelská scénářová položka s minimálním prahem 3 000 obyvatel; nenahrazuje maloobchodní prodejnu."
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
const RTS_811_5_CONSTRUCTION_RATES = {
  "1": 8995,
  "2": 10450,
  "3": 8460,
  "4": 8080
};
const RTS_801_GROUP_BY_ITEM_ID = {
  VYB_01_J: "801.3", VYB_01: "801.3", VYB_02A: "801.3", VYB_02B: "801.3",
  VYB_02C_1: "801.3", VYB_02C_2: "801.3",
  VYB_05B: "801.5",
  VYB_06A: "801.9", VYB_06B: "801.9", VYB_07A_1: "801.9", VYB_07A_2: "801.9",
  VYB_09B_1: "801.1", VYB_09B_2: "801.1", VYB_09A_4: "801.1", VYB_16: "801.1",
  VYB_17: "801.1", VYB_18: "801.1", VYB_19: "801.1", VYB_20: "801.1",
  VYB_09D: "801.6", VYB_16A: "801.6", [VOLUNTEER_FIRE_STATION_ID]: "801.6", VYB_17A: "801.6",
  VYB_10A: "801.4", VYB_10B_1: "801.4", VYB_10B_2: "801.4",
  VYB_11: "801.6",
  VYB_13A: "801.8", VYB_13B: "801.8",
  VYB_15A_1: "811.5", VYB_15A_1_LAW: "811.5",
  VYB_15A_1_RESERVE: "811.5", VYB_15A_2: "811.5", VYB_15A_3: "811.5"
};
const RTS_801_SOURCE = "RTS, Cenové ukazatele ve stavebnictví pro rok 2026, s. 1, skupina JKSO 801; ceny v Kč/m³ obestavěného prostoru (m³ OP).";
const RTS_811_5_SOURCE = "RTS, Cenové ukazatele ve stavebnictví pro rok 2026, s. 1, JKSO 811.5 (haly pro garážování, opravy a údržbu vozidel, strojů a zařízení); ceny v Kč/m³ obestavěného prostoru (m³ OP).";
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
      { code: "S_uzemi" }, { code: "S_zastav" }, { code: "TYP_UZEMI" },
      { code: "KOD_TVAR" }, { field: COVERAGE_OVERLAP_INPUT, secondary: true },
      { code: "Plocha pro odhad dostupnosti", secondary: true }
    ]
  },
  {
    label: "Počty obyvatel",
    fields: [
      { code: "O_sidlo" }, { code: "O_celk" }, { code: "O_spad" }
    ]
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
      { code: "PARKING_LARGE_APARTMENT_SHARE" },
      { code: "PARKING_SMALL_APARTMENT_AREA_M2" },
      { code: "PARKING_LARGE_APARTMENT_AREA_M2" },
      { code: "PARKING_FAMILY_HOME_AREA_M2" },
      { code: "KOMPAKT" },
      { code: "Průměrná podlažnost", secondary: true },
      { code: "HPP na byt", secondary: true },
      { code: "Koeficient zastavění", secondary: true }
    ]
  }
];
const GENERAL_OBLIGATION_RULE_IDS = new Set(["ZEL-LEG-002", "SAFE-LEG-001", "TECH-LEGAL-002"]);
const HIDDEN_DUPLICATE_ITEM_IDS = new Set(["VYB_16B"]);
const REMOVED_SOCIAL_SERVICE_ITEM_IDS = new Set(["VYB_06C", "VYB_07A_1", "VYB_07A_2"]);
const RULE_MODEL_LINKS = {
  "ZEL-LEG-001": ["VYB_04A"],
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
  "ZDR-MET-001": ["VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20", "VYB_09A_4"],
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
  "SAFE-MET-001": ["VYB_16B", VOLUNTEER_FIRE_STATION_ID],
  "SAFE-LEG-001": ["VYB_17B"],
  "TECH-LEGAL-002": ["VYB_18A", "VYB_18B"],
  "TECH-MET-001": ["VYB_18A"],
  "TECH-MET-002": ["VYB_18B"],
  "TECH-MET-003": ["VYB_19"],
  "TECH-HYDRO-001": ["VYB_20A"],
  "TECH-COST-001": ["VYB_20B_1", "VYB_20B_2", "VYB_20B_3"],
  "TECH-MET-005": ["VYB_18B", "VYB_19"],
};
function createMmrRecommendationRule(id, title, description, condition, page) {
  return {
    [RULE_HEADERS.id]: id,
    [RULE_HEADERS.equipment]: title,
    [RULE_HEADERS.type]: "Doporučený standard dostupnosti",
    [RULE_HEADERS.description]: description,
    [RULE_HEADERS.condition]: condition,
    [RULE_HEADERS.verification]: "Jde o metodické doporučení MMR, nikoli samostatnou zákonnou povinnost.",
    [RULE_HEADERS.limitation]: "Ověřte skutečnou polohu, kapacitu, trasu a místní podmínky; model nenahrazuje GIS ani vyjádření poskytovatele.",
    [RULE_HEADERS.source]: "TB050MMR01 – Standardy dostupnosti veřejné infrastruktury, aktualizace 2020",
    [RULE_HEADERS.location]: page
  };
}
const MMR_PUBLIC_SPACE_RULE = createMmrRecommendationRule(
  "ZEL-MMR-PARK-001",
  "Veřejná parková zeleň – doporučená dostupnost MMR",
  "Parková zeleň jako součást veřejného prostranství: doporučená plocha 0,5–1 ha; typy A/B/C do 300 m skutečné pěší vzdálenosti, typ D v každém obytném sídle.",
  "Počet bodů pro typy A/B/C se odhaduje z 300m sítě; pro typ D určete počet obytných sídel samostatně.",
  "Tabulka 4, tištěná str. 17"
);
const MMR_SPECIALIST_RULES = {
  VYB_09B_1: createMmrRecommendationRule(
    "ZDR-MMR-GROUP-2",
    "Ambulantní specialisté – skupina 2 MMR",
    "Metodika doporučuje sídelně-strukturální přítomnost péče skupiny 2 v obcích typu A; v typu B nad 5 000 obyvatel doporučuje přítomnost. Nejde o samostatnou populační kvótu ordinací.",
    "Posuďte podle typu území, hranice populace a skutečné dostupnosti; tento standard neurčuje zákonnou povinnost zřídit místní ambulanci.",
    "Tabulka 1, tištěná str. 10"
  ),
  VYB_09B_2: createMmrRecommendationRule(
    "ZDR-MMR-GROUP-3",
    "Ambulantní specialisté – skupina 3 MMR",
    "Metodika doporučuje sídelně-strukturální přítomnost péče skupiny 3 v obcích typu A; v typu B nad 5 000 obyvatel doporučuje přítomnost. Nejde o samostatnou populační kvótu ordinací.",
    "Posuďte podle typu území, hranice populace a skutečné dostupnosti; tento standard neurčuje zákonnou povinnost zřídit místní ambulanci.",
    "Tabulka 1, tištěná str. 10"
  )
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
const MMR_SCHOOL_RECOMMENDED_ITEM_IDS = new Set(SCHOOL_CAPACITY_ITEM_IDS);
const MMR_SCHOOL_STANDARDS = {
  VYB_01: {
    populationPerSeat: 28,
    minimumPopulation: 420,
    minimumCapacity: 15,
    name: "mateřskou školu",
    densityByDistance: { 400: 14, 600: 6.3 }
  },
  VYB_02A: {
    populationPerSeat: 22,
    minimumPopulation: 2000,
    minimumCapacity: 85,
    smallSchoolPopulation: 220,
    smallSchoolCapacity: 10,
    name: "ZŠ 1. stupně",
    densityByDistance: { 600: 30, 800: 16.8 },
    capacityTiers: [
      { population: 2000, minimumSeats: 85, maximumSeats: 150 },
      { population: 10000, minimumSeats: 450, maximumSeats: 450 }
    ]
  },
  VYB_02B: {
    populationPerSeat: 29,
    minimumPopulation: 4000,
    minimumCapacity: 136,
    maximumCapacity: 360,
    name: "úplnou základní školu",
    densityByDistance: { 800: 33.6 },
    capacityTiers: [
      { population: 4000, minimumSeats: 136, maximumSeats: 360 },
      { population: 10000, minimumSeats: 360, maximumSeats: 360 }
    ]
  }
};
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
  VYB_01: { ruleId: "MS-MET-001", distanceMeters: "school-access-rule" },
  VYB_02A: { ruleId: "ZS-MET-001", distanceMeters: "school-access-rule" },
  VYB_02B: { ruleId: "ZS-MET-001", distanceMeters: "school-access-rule" },
  VYB_03A: { ruleId: "ZEL-MET-001", distanceMeters: 200 },
  VYB_03B: { ruleId: "ZEL-MET-002", distanceMeters: 500 },
  VYB_03C: { ruleId: "ZEL-MET-003", distanceMeters: "by-settlement-type" },
  VYB_04B: { ruleId: "ZEL-MMR-PARK-001", distanceMeters: "by-settlement-type" },
  VYB_06A: { ruleId: "SOC-MET-001", distanceMeters: "day-service" },
  VYB_14: { ruleId: "MOB-MET-001", distanceMeters: "public-transport-stop" },
  VYB_09A_4: { ruleId: "ZDR-MET-001", distanceMeters: "healthcare-group-1" },
  VYB_10A: { ruleId: "CULT-MET-002", distanceMeters: "community-center" },
  VYB_16: { ruleId: "ZDR-MET-001", distanceMeters: "healthcare-group-1" },
  VYB_17: { ruleId: "ZDR-MET-001", distanceMeters: "healthcare-group-1" },
  VYB_18: { ruleId: "ZDR-MET-001", distanceMeters: "healthcare-group-1" },
  VYB_20: { ruleId: "ZDR-MET-001", distanceMeters: "healthcare-group-1" },
  VYB_19: [
    { ruleId: "ZDR-MET-001", distanceMeters: "healthcare-group-1", applyPointFloor: true },
    { ruleId: "TECH-MET-003", distanceMeters: "recycling-yard", applyPointFloor: true }
  ],
  VYB_16B: {
    ruleId: "SAFE-MET-001",
    distanceMeters: 1000,
    applyPointFloor: false,
    pointFloorNote: "Pěší dosah zbrojnice je pouze jedna z kontrol. Potřebu nové zbrojnice nelze automaticky odvodit bez posouzení JPO, akceschopnosti a požárního poplachového plánu."
  },
  [VOLUNTEER_FIRE_STATION_ID]: {
    ruleId: "SAFE-MET-001",
    distanceMeters: 1000,
    applyPointFloor: false
  },
  VYB_18A: { ruleId: "TECH-MET-001", distanceMeters: 50 },
  VYB_18B: { ruleId: "TECH-MET-002", distanceMeters: "recycling-collection" }
};
const SPATIAL_POINT_FLOOR_ITEM_IDS = new Set([
  "VYB_03A", "VYB_03B", "VYB_04B",
  "VYB_10A", "VYB_14", "VYB_18A", "VYB_18B",
  "VYB_09A_4", "VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20"
]);
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
  VYB_03C: "Hřiště pro mládež a dospělé",
  VYB_18A: "Stanoviště SKO a separace",
  VYB_18B: "Podzemní/polopodzemní kontejnerová hnízda",
  VYB_06A: "Centrum denních služeb a denní stacionář",
  VYB_06B: "Domov pro seniory, dům s pečovatelskou službou",
  VYB_04A: "Veřejné prostranství",
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
const SOCIAL_SERVICE_AVAILABILITY_STANDARDS = [
  {
    key: "day-services",
    label: "Centrum denních služeb a denní stacionář",
    details: "Doporučený standard: Typ A/B/C – skutečná pěší docházka do 600 m; typ D – dojezdová doba do 30 minut.",
    territories: ["A", "B", "C", "D"],
    accessStandard: true,
    thresholdText: "Typ A/B/C: skutečná pěší docházka do 600 m; typ D: dojezdová doba do 30 minut",
    itemId: "VYB_06A",
    demographic: {
      coefficient: 0.004,
      unitLabel: "míst",
      unitForms: ["místo", "místa", "míst"],
      usableAreaPerUnit: 14,
      siteAreaPerUnit: 20,
      source: "Demografický scénář: 20 % obyvatel v cílové seniorské nebo obdobně potřebné skupině × 2 % z nich potřebujících současně denní službu = 0,004 místa na obyvatele. Jde o návrhový předpoklad, nikoli statistický normativ; upravte podle věkové struktury a poskytovatele."
    }
  },
  {
    key: "low-threshold-day-center",
    label: "Nízkoprahové denní centrum",
    details: "Samostatná ambulantní služba. Doporučený standard: Typ A/B/C, obec nad 5 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    territories: ["A", "B", "C"],
    minimumPopulation: 5000,
    thresholdText: "Typ A/B/C, obce > 5 000 obyvatel: sídelně-strukturální přítomnost v obci",
    templateId: "VYB_06A",
    demographic: {
      coefficient: 0.001,
      unitLabel: "míst",
      unitForms: ["místo", "místa", "míst"],
      usableAreaPerUnit: 14,
      siteAreaPerUnit: 20,
      source: "Demografický scénář: 1 % obyvatel v cílové skupině × 10 % současně využívajících službu = 0,001 místa na obyvatele. Podíl cílové skupiny i souběžné využití jsou pracovní předpoklady, nikoli statistický normativ."
    }
  },
  {
    key: "ambulatory-social-care",
    label: "Ostatní zařízení ambulantní sociální péče",
    details: "Doporučený standard MMR se vztahuje na nízkoprahové denní centrum, nízkoprahové zařízení pro děti a mládež, týdenní stacionář a intervenční centrum: Typ A/B/C, obce > 5 000 obyvatel – sídelně-strukturální přítomnost v obci. Kontaktní centra a zařízení sociálně aktivizačních služeb jsou v modelu zahrnuta, ale tabulka 1 jim tento práh nepřiřazuje samostatně.",
    territories: ["A", "B", "C"],
    minimumPopulation: 5000,
    thresholdText: "Typ A/B/C, obce > 5 000 obyvatel: sídelně-strukturální přítomnost v obci",
    templateId: "VYB_06A",
    demographic: {
      coefficient: 0.002,
      unitLabel: "míst",
      unitForms: ["místo", "místa", "míst"],
      usableAreaPerUnit: 14,
      siteAreaPerUnit: 20,
      source: "Skupinový demografický scénář pro více různých služeb: 5 % obyvatel s předpokládanou potřebou podpory × 4 % současně využívajících některou z těchto služeb = 0,002 místa na obyvatele. Jde jen o společný pracovní odhad; ověřte skladbu služeb a místní poptávku."
    }
  },
  {
    key: "night-shelter",
    label: "Noclehárna",
    details: "Doporučený standard: Typ A/B, obce > 30 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    territories: ["A", "B"],
    minimumPopulation: 30000,
    thresholdText: "Typ A/B, obce > 30 000 obyvatel: sídelně-strukturální přítomnost v obci",
    templateId: "VYB_07A_1",
    demographic: {
      coefficient: 0.0005,
      unitLabel: "lůžek",
      unitForms: ["lůžko", "lůžka", "lůžek"],
      usableAreaPerUnit: 32,
      siteAreaPerUnit: 40,
      source: "Demografický scénář: 0,5 % obyvatel v modelované skupině osob bez stabilního bydlení × 10 % současně potřebujících lůžko = 0,0005 lůžka na obyvatele. Jde o pracovní předpoklad, nikoli normativ; podklad ÚÚR pro noclehárnu přímo odpovídající není."
    }
  },
  {
    key: "crisis-services",
    label: "Azylový dům a zařízení pro krizovou pomoc",
    details: "Doporučený standard: Typ A/B, obce > 30 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    territories: ["A", "B"],
    minimumPopulation: 30000,
    thresholdText: "Typ A/B, obce > 30 000 obyvatel: sídelně-strukturální přítomnost v obci",
    templateId: "VYB_07A_1",
    demographic: {
      coefficient: 0.0005,
      unitLabel: "lůžek",
      unitForms: ["lůžko", "lůžka", "lůžek"],
      usableAreaPerUnit: 35,
      siteAreaPerUnit: 50,
      source: "Skupinový demografický scénář: 1 % obyvatel v modelované skupině ohrožené bytovou či jinou krizí × 5 % současně potřebujících lůžko = 0,0005 lůžka na obyvatele. Jde o pracovní předpoklad bez společného normativu; ověřte skladbu služeb a místní potřebu."
    }
  },
  {
    key: "halfway-house",
    label: "Dům na půl cesty",
    details: "Doporučený standard: Typ A/B, obce > 50 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    territories: ["A", "B"],
    minimumPopulation: 50000,
    thresholdText: "Typ A/B, obce > 50 000 obyvatel: sídelně-strukturální přítomnost v obci",
    templateId: "VYB_07A_1",
    demographic: {
      coefficient: 0.0001,
      unitLabel: "lůžek",
      unitForms: ["lůžko", "lůžka", "lůžek"],
      usableAreaPerUnit: 35,
      siteAreaPerUnit: 50,
      source: "Demografický scénář: 0,2 % obyvatel v modelované skupině mladých lidí potřebujících návazné bydlení × 5 % současně umístěných = 0,0001 lůžka na obyvatele. Jde o pracovní předpoklad, nikoli normativ; ověřte místní potřebu a spád."
    }
  },
  {
    key: "care-service-housing",
    label: "Domov pro seniory, dům s pečovatelskou službou",
    details: "Doporučený standard MMR pro domov pro seniory: Typ A – sídelně-strukturální přítomnost v obci. Dům s pečovatelskou službou je zahrnut do společné modelové položky; tabulka 1 mu samostatný práh nestanoví.",
    territories: ["A"],
    thresholdText: "Domov pro seniory: Typ A, sídelně-strukturální přítomnost v obci; pro dům s pečovatelskou službou tabulka 1 samostatný práh neuvádí",
    itemId: "VYB_06B",
    demographic: {
      coefficient: 0.01,
      unitLabel: "bytů",
      unitForms: ["byt", "byty", "bytů"],
      usableAreaPerUnit: 45,
      siteAreaPerUnit: 60,
      source: "Demografický scénář položky Domov pro seniory, dům s pečovatelskou službou: 10 % obyvatel v seniorské skupině × 10 % z nich potřebujících byt s pečovatelskou službou = 0,01 bytu na obyvatele. Podíly jsou návrhové předpoklady k místnímu ověření, nikoli statistický normativ. ÚÚR uvádí odděleně orientační výměru 45 m² podlažní plochy a 60 m² pozemku na byt."
    }
  },
  {
    key: "disability-services",
    label: "Domov pro osoby se zdravotním postižením a chráněné bydlení",
    details: "Doporučený standard: Typ A/B, obce > 30 000 obyvatel – sídelně-strukturální přítomnost v obci. Položka zahrnuje domov pro osoby se zdravotním postižením a chráněné bydlení.",
    territories: ["A", "B"],
    minimumPopulation: 30000,
    thresholdText: "Typ A/B, obce > 30 000 obyvatel: sídelně-strukturální přítomnost v obci",
    templateId: "VYB_07A_1",
    demographic: {
      coefficient: 0.002,
      unitLabel: "míst",
      unitForms: ["místo", "místa", "míst"],
      usableAreaPerUnit: 40,
      siteAreaPerUnit: 60,
      source: "Skupinový demografický scénář: 10 % obyvatel s předpokládaným zdravotním postižením × 2 % potřebujících pobytovou nebo chráněnou formu podpory = 0,002 místa na obyvatele. Jde o návrhové podíly, nikoli statistický normativ; společný výsledek pro domov i chráněné bydlení upravte podle místních údajů."
    }
  }
];

function addSocialAvailabilityCatalogItems(items) {
  const category = "SOCIÁLNÍ SLUŽBY A DOSTUPNÉ BYDLENÍ";
  const createRule = (standard) => ({
    [RULE_HEADERS.id]: `SOC-MMR-${standard.key.toUpperCase()}`,
    [RULE_HEADERS.equipmentId]: `SOCIAL_${standard.key.toUpperCase()}`,
    [RULE_HEADERS.equipment]: standard.label,
    [RULE_HEADERS.area]: "Sociální služby",
    [RULE_HEADERS.type]: "Doporučený standard dostupnosti",
    [RULE_HEADERS.condition]: standard.thresholdText,
    [RULE_HEADERS.description]: standard.details || standard.thresholdText,
    [RULE_HEADERS.formula]: "Posouzení typu území, populačního prahu a dostupnosti služby.",
    [RULE_HEADERS.inputs]: "Typ území A/B/C/D; počet obyvatel řešeného území po záměru a spádu; ověřená přítomnost služby.",
    [RULE_HEADERS.output]: "Doporučení k ověření sídelní dostupnosti; kapacita a cena se počítají samostatně z položkového scénáře.",
    [RULE_HEADERS.rounding]: "Bez zaokrouhlování.",
    [RULE_HEADERS.source]: "Výpis standardů sociálních služeb z podkladů projektu.",
    [RULE_HEADERS.location]: standard.thresholdText,
    [RULE_HEADERS.url]: "",
    [RULE_HEADERS.verification]: "Převzato z podkladů projektu; primární publikace a aktuální znění je nutné ověřit.",
    [RULE_HEADERS.limitation]: "Doporučení přítomnosti není samo o sobě povinností výstavby ani kapacitní kvótou."
  });
  for (const standard of SOCIAL_SERVICE_AVAILABILITY_STANDARDS) {
    if (standard.itemId) {
      const item = items.find((candidate) => candidate.id === standard.itemId);
      if (!item) throw new Error(`Chybí katalogová položka ${standard.itemId} pro ${standard.label}.`);
      item.category = category;
      item.socialAvailabilityStandard = {
        ...standard,
        templateName: displayCatalogItemName(item)
      };
      item.socialDemographic = standard.demographic;
      continue;
    }
    const id = `SOCIAL_${standard.key.toUpperCase().replace(/-/g, "_")}`;
    const template = items.find((item) => item.id === standard.templateId);
    if (!template) {
      throw new Error(`Chybí katalogová předloha ${standard.templateId} pro ${standard.label}.`);
    }
    items.push({
      ...template,
      key: `${id}#1`,
      id,
      name: standard.label,
      category,
      legalRules: [],
      recommendedRules: [createRule(standard)],
      otherRules: [],
      sourceFormulas: { ...template.sourceFormulas },
      metricSources: { ...template.metricSources },
      userAdded: false,
      socialTemplateName: displayCatalogItemName(template),
      socialAvailabilityStandard: {
        ...standard,
        templateName: displayCatalogItemName(template)
      },
      socialDemographic: standard.demographic
    });
  }
}

function addVolunteerFireStationItem(items) {
  if (items.some((item) => item.id === VOLUNTEER_FIRE_STATION_ID)) return;
  const template = items.find((item) => item.id === "VYB_16A");
  if (!template) throw new Error("Chybí položka HZS jako předloha pro zbrojnici JSDH.");
  items.push({
    ...template,
    key: `${VOLUNTEER_FIRE_STATION_ID}#1`,
    id: VOLUNTEER_FIRE_STATION_ID,
    name: "Hasičská zbrojnice dobrovolných hasičů (JSDH)",
    sourceRow: null,
    sourceFormulas: {},
    metricSources: {},
    legalRules: [],
    recommendedRules: [],
    otherRules: [],
    manualCoverageCheck: false
  });
}

const UUR_SOCIAL_AREA_REFERENCES = [
  {
    label: "Centrum denních služeb / denní stacionář",
    unitLabel: "míst profesionální pečovatelky",
    unitsPerThousand: 0.4,
    floorAreaPerUnit: 14,
    siteAreaPerUnit: 20
  },
  {
    label: "Domov pro seniory, dům s pečovatelskou službou",
    unitLabel: "bytů",
    unitsPerThousand: 2,
    floorAreaPerUnit: 45,
    siteAreaPerUnit: 60
  },
  {
    label: "Domov pro osoby se zdravotním postižením",
    unitLabel: "míst",
    unitsPerThousand: 1.2,
    floorAreaPerUnit: 40,
    siteAreaPerUnit: 60
  },
  {
    label: "Ústav sociální péče pro děti a mládež (všech druhů postižení)",
    unitLabel: "míst",
    unitsPerThousand: 1.8,
    floorAreaPerUnit: 45,
    siteAreaPerUnit: 60
  },
  {
    label: "Zařízení pro denní pobyt mládeže",
    unitLabel: "míst",
    unitsPerThousand: 0.1,
    floorAreaPerUnit: 30,
    siteAreaPerUnit: 50
  },
  {
    label: "Azylový dům pro matky/otce s dětmi",
    unitLabel: "míst pro rodiče",
    unitsPerThousand: 0.1,
    floorAreaPerUnit: 50,
    siteAreaPerUnit: 60
  },
  {
    label: "Ubytovna pro přechodný pobyt (orientační analogie, ne noclehárna)",
    unitLabel: "lůžek",
    unitsPerThousand: 0.05,
    floorAreaPerUnit: 32,
    siteAreaPerUnit: 40
  }
];
const HEALTHCARE_DEMAND_ITEM_IDS = new Set([
  "VYB_09A_4", "VYB_09B_1", "VYB_09B_2", "VYB_17", "VYB_18", "VYB_19"
]);
const MMR_HEALTHCARE_GROUP_1_ITEM_IDS = new Set([
  "VYB_09A_4", "VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20"
]);
const HEALTHCARE_ACCESS_INPUTS = [
  { code: "HEALTHCARE_WALK_DISTANCE", label: "Ověřená pěší vzdálenost k ordinaci (A/B/C; limit 600 m)", kind: "number", min: 0, step: 1, unit: "m" },
  { code: "HEALTHCARE_DRIVE_MINUTES", label: "Ověřený dojezd k ordinaci (D; limit 35 min)", kind: "number", min: 0, step: 0.1, unit: "min" },
  { code: "HEALTHCARE_ACCESS_SOURCE", label: "Zdroj / datum ověření dostupnosti", kind: "text", unit: "" }
];
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
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    { code: "EARLY_CHILDCARE_EXISTING_CAPACITY", label: "Ověřená volná kapacita okolních zařízení", kind: "number", min: 0, step: 1, unit: "míst" },
    { code: "EARLY_CHILDCARE_ACCESS_DISTANCE", label: "Skutečná pěší vzdálenost k zařízení s volnou kapacitou", kind: "number", min: 0, step: 1, unit: "m" },
    { code: "EARLY_CHILDCARE_DATA_SOURCE", label: "Zdroj demografie a kapacit", kind: "text", unit: "" }
  ],
  VYB_09A_4: [
    { code: "MODEL_DEMOGRAPHIC_COEFFICIENT", label: "Podíl přírůstku obyvatel započtený do poptávky", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "Výchozí výpočet zohledňuje celý přírůstek obyvatel; podíl je editovatelný." },
    { code: "MODEL_LEGAL_PATIENTS_PER_DOCTOR", label: "Registrovaní pacienti na 1 lékaře při plném úvazku (státní pracovní základ)", kind: "number", min: 1, step: 50, defaultValue: 1700, unit: "pacientů / lékaře na plný úvazek", note: "Výchozí předpoklad 1 700 pacientů (střed pracovního rozmezí 1 600–1 800). Zákon neurčuje jednotnou kvótu; údaj je editovatelný zvlášť pro tuto položku." },
    { code: "MODEL_POPULATION_PER_FTE", label: "Optimální počet obyvatel na lékaře (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 14000, unit: "obyvatel/lékaře", note: "MMR, s. 27: 14 000 obyvatel na lékaře dle optimálního zajištění VZP. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinační dobu 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 5600, unit: "obyvatel/lékaře", note: "MMR, s. 27: 5 600 obyvatel při ordinaci 2 dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 84, unit: "obyvatel/ha", note: "MMR, s. 27: 84 obyvatel/ha pro gynekologii při ordinaci 2 dny týdně; metodika upozorňuje, že prah nelze mechanicky aplikovat bez spádovosti." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné ordinace", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "lékařského úvazku/ordinaci" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 120, unit: "m²/úvazek" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 360, unit: "m²/úvazek" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
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
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
  ],
  VYB_02C_2: [
    { code: "MODEL_DEMOGRAPHIC_SHARE", label: "Modelový podíl nových žáků (analogicky ZŠ I. stupně)", kind: "percent", min: 0, max: 1, step: 0.001, defaultValue: 1 / 22, unit: "%", note: "Scénářový přepočet podle MMR poměru ZŠ I. stupně 1 žák / 22 obyvatel; MMR kvótu žáků ZUŠ nestanoví." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Minimální modelová kapacita jedné ZUŠ (analogicky ZŠ I. stupně)", kind: "number", min: 1, step: 1, defaultValue: 85, unit: "žákovských míst", note: "Modelový přenos minima 85 míst pro ZŠ I. stupně podle MMR; nejde o vlastní kapacitní kvótu MMR pro ZUŠ." },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na žákovské místo", kind: "number", min: 0, step: 0.1, defaultValue: 8, unit: "m²/místo" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na žákovské místo", kind: "number", min: 0, step: 0.1, defaultValue: 12, unit: "m²/místo" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0, step: 0.01, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.4 (kultura a osvěta; proxy)", kind: "number", min: 0, step: 1, defaultValue: 14760, unit: "Kč/m³", note: "Cenový proxy pro vzdělávání v ZUŠ; upravte podle konkrétního stavebního programu." },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
  ],
  VYB_17: [
    { code: "MODEL_DEMOGRAPHIC_COEFFICIENT", label: "Demografický koeficient započtené populace", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "Podíl obyvatel, ze kterého se odvozuje populační spád ordinace." },
    { code: "MODEL_LEGAL_PATIENTS_PER_DOCTOR", label: "Registrovaní pacienti na 1 lékaře při plném úvazku (státní pracovní základ)", kind: "number", min: 1, step: 50, defaultValue: 1700, unit: "pacientů / lékaře na plný úvazek", note: "Výchozí předpoklad 1 700 pacientů (střed pracovního rozmezí 1 600–1 800). Zákon neurčuje jednotnou kvótu; údaj je editovatelný zvlášť pro tuto položku." },
    { code: "MODEL_POPULATION_PER_FTE", label: "Optimální počet obyvatel na lékaře (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 1800, unit: "obyvatel/lékaře", note: "MMR, s. 27: 1 800 obyvatel na lékaře dle optimálního zajištění VZP. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 10, defaultValue: 720, unit: "obyvatel/lékaře", note: "MMR, s. 27: 720 obyvatel při ordinaci dva dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 11, unit: "obyvatel/ha", note: "MMR, s. 27: 11 obyvatel/ha. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinační dobu 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 720, unit: "obyvatel/lékaře", note: "MMR, s. 27: 720 obyvatel při ordinaci 2 dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 11, unit: "obyvatel/ha", note: "MMR, s. 27: 11 obyvatel/ha při ordinaci 2 dny týdně; plný standard je 27 obyvatel/ha." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné ordinace", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "lékařského úvazku/ordinaci" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 45, unit: "m²/úvazek" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 135, unit: "m²/úvazek" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0.1, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    ...HEALTHCARE_ACCESS_INPUTS
  ],
  VYB_18: [
    { code: "MODEL_DEMOGRAPHIC_COEFFICIENT", label: "Podíl nové populace započtené do poptávky", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "MMR uvádí počet obyvatel na lékaře, nikoli počet dětských pacientů; výpočet používá celou populaci." },
    { code: "MODEL_LEGAL_PATIENTS_PER_DOCTOR", label: "Registrovaní pacienti na 1 lékaře při plném úvazku (státní pracovní základ)", kind: "number", min: 1, step: 50, defaultValue: 1700, unit: "pacientů / lékaře na plný úvazek", note: "Výchozí předpoklad 1 700 pacientů (střed pracovního rozmezí 1 600–1 800). Zákon neurčuje jednotnou kvótu; údaj je editovatelný zvlášť pro tuto položku." },
    { code: "MODEL_POPULATION_PER_FTE", label: "Optimální počet obyvatel na lékaře (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 4500, unit: "obyvatel/lékaře", note: "MMR, s. 27: 4 500 obyvatel na lékaře dle optimálního zajištění VZP. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 10, defaultValue: 1800, unit: "obyvatel/lékaře", note: "MMR, s. 27: 1 800 obyvatel při ordinaci dva dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 27, unit: "obyvatel/ha", note: "MMR, s. 27: 27 obyvatel/ha. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinační dobu 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 1800, unit: "obyvatel/lékaře", note: "MMR, s. 27: 1 800 obyvatel při ordinaci 2 dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 27, unit: "obyvatel/ha", note: "MMR, s. 27: 27 obyvatel/ha při ordinaci 2 dny týdně; plný standard je 68 obyvatel/ha." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné ordinace", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "lékařského úvazku/ordinaci" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 55, unit: "m²/úvazek" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 165, unit: "m²/úvazek" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0.1, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    ...HEALTHCARE_ACCESS_INPUTS
  ],
  VYB_19: [
    { code: "MODEL_DEMOGRAPHIC_COEFFICIENT", label: "Demografický koeficient započtené populace", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "Podíl obyvatel, ze kterého se odvozuje populační spád ordinace." },
    { code: "MODEL_LEGAL_PATIENTS_PER_DOCTOR", label: "Registrovaní pacienti na 1 lékaře při plném úvazku (státní pracovní základ)", kind: "number", min: 1, step: 50, defaultValue: 1700, unit: "pacientů / lékaře na plný úvazek", note: "Výchozí předpoklad 1 700 pacientů (střed pracovního rozmezí 1 600–1 800). Zákon neurčuje jednotnou kvótu; údaj je editovatelný zvlášť pro tuto položku." },
    { code: "MODEL_POPULATION_PER_FTE", label: "Optimální počet obyvatel na lékaře (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 2100, unit: "obyvatel/lékaře", note: "MMR, s. 27: 2 100 obyvatel na lékaře dle optimálního zajištění VZP. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 10, defaultValue: 840, unit: "obyvatel/lékaře", note: "MMR, s. 27: 840 obyvatel při ordinaci dva dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 13, unit: "obyvatel/ha", note: "MMR, s. 27: 13 obyvatel/ha. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_POPULATION_PER_FTE", label: "Minimální počet obyvatel pro ordinační dobu 2 dny týdně (MMR 2020)", kind: "number", min: 1, step: 100, defaultValue: 840, unit: "obyvatel/lékaře", note: "MMR, s. 27: 840 obyvatel při ordinaci 2 dny týdně. Editovatelný metodický vstup." },
    { code: "MODEL_MINIMUM_DENSITY_PER_HA", label: "Minimální hustota pro ordinaci 2 dny týdně (MMR 2020)", kind: "number", min: 0, step: 1, defaultValue: 13, unit: "obyvatel/ha", note: "MMR, s. 27: 13 obyvatel/ha při ordinaci 2 dny týdně; plný standard je 32 obyvatel/ha." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné ordinace", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "lékařského úvazku/ordinaci" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 45, unit: "m²/úvazek" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na lékařský úvazek", kind: "number", min: 0, step: 1, defaultValue: 135, unit: "m²/úvazek" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0.1, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    ...HEALTHCARE_ACCESS_INPUTS
  ],
  VYB_09B_1: [
    { code: "MODEL_DEMOGRAPHIC_COEFFICIENT", label: "Demografický koeficient spádové populace", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "Podíl obyvatel zahrnutý do populačního odhadu; místní oborová skladba může vyžadovat úpravu." },
    { code: "MODEL_LEGAL_PATIENTS_PER_DOCTOR", label: "Registrovaní pacienti na 1 lékaře při plném úvazku (státní pracovní základ)", kind: "number", min: 1, step: 50, defaultValue: 1700, unit: "pacientů / lékaře na plný úvazek", note: "Výchozí předpoklad 1 700 pacientů (střed pracovního rozmezí 1 600–1 800). Zákon neurčuje jednotnou kvótu; údaj je editovatelný zvlášť pro tuto skupinu odborností." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné odborné ambulance", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "ambulance" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na jednu ambulanci", kind: "number", min: 0, step: 1, defaultValue: 70, unit: "m²/ambulanci" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jednu ambulanci", kind: "number", min: 0, step: 1, defaultValue: 210, unit: "m²/ambulanci" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0.1, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" }
  ],
  VYB_09B_2: [
    { code: "MODEL_DEMOGRAPHIC_COEFFICIENT", label: "Demografický koeficient spádové populace", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 1, unit: "%", note: "Podíl obyvatel zahrnutý do populačního odhadu; místní oborová skladba může vyžadovat úpravu." },
    { code: "MODEL_LEGAL_PATIENTS_PER_DOCTOR", label: "Registrovaní pacienti na 1 lékaře při plném úvazku (státní pracovní základ)", kind: "number", min: 1, step: 50, defaultValue: 1700, unit: "pacientů / lékaře na plný úvazek", note: "Výchozí předpoklad 1 700 pacientů (střed pracovního rozmezí 1 600–1 800). Zákon neurčuje jednotnou kvótu; údaj je editovatelný zvlášť pro tuto skupinu odborností." },
    { code: "MODEL_CAPACITY_PER_FACILITY", label: "Kapacita jedné odborné ambulance", kind: "number", min: 0.1, step: 0.1, defaultValue: 1, unit: "ambulance" },
    { code: "MODEL_AREA_PER_UNIT", label: "Užitná plocha na jednu ambulanci", kind: "number", min: 0, step: 1, defaultValue: 70, unit: "m²/ambulanci" },
    { code: "MODEL_SITE_PER_UNIT", label: "Plocha areálu na jednu ambulanci", kind: "number", min: 0, step: 1, defaultValue: 210, unit: "m²/ambulanci" },
    { code: "MODEL_FLOOR_FACTOR", label: "Koeficient celkové podlahové plochy", kind: "number", min: 0.1, step: 0.05, defaultValue: 1.1, unit: "m²/m²" },
    { code: "MODEL_VOLUME_PER_FLOOR_M2", label: "Obestavěný prostor na m² podlahové plochy", kind: "number", min: 0.1, step: 0.1, defaultValue: 3.3, unit: "m³/m²" },
    { code: "MODEL_PAVED_SHARE", label: "Podíl zpevněné plochy areálu", kind: "percent", min: 0, max: 1, step: 0.01, defaultValue: 0.25, unit: "%" },
    { code: "MODEL_COST_PER_M3", label: "RTS 2026 – JKSO 801.1 (zdravotnické stavby)", kind: "number", min: 0, step: 1, defaultValue: 13645, unit: "Kč/m³" },
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" }
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
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    { code: "ZZS_COVERED_POPULATION", label: "Obyvatelé projektu a spádu dosažitelní ze stávající základny do 20 min", kind: "number", min: 0, step: 1, unit: "obyvatel" },
    { code: "ZZS_MAX_RESPONSE_MINUTES", label: "Ověřený dojezd ZZS (zákonná mez 20 min)", kind: "number", min: 0, step: 0.1, unit: "min", note: "Zákon č. 374/2011 Sb., § 5: nejpozději do 20 minut od převzetí pokynu k výjezdu, s výjimkami stanovenými zákonem." },
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
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    { code: "HZS_COVERED_POPULATION", label: "Obyvatelé spádu pokrytí stávající jednotkou", kind: "number", min: 0, step: 1, unit: "obyvatel" },
    { code: "HZS_RESPONSE_MINUTES", label: "Ověřený zásahový dojezd stanice (A/B ≤ 7 min; C/D ≤ 15 min)", kind: "number", min: 0, step: 0.1, unit: "min", note: "Zadejte skutečný dojezd podle požárního poplachového plánu; jde o silniční zásahový dojezd, nikoli pěší vzdálenost." },
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
    { code: "MODEL_COST_SURCHARGE", label: "VRN a rozpočtová rezerva", kind: "percent", min: 0, max: 3, step: 0.01, defaultValue: 0.08, unit: "%" },
    { code: "POLICE_RESPONSE_MINUTES", label: "Ověřená doba zásahu hlídky (A/B ≤ 10 min; C ≤ 15 min; D se neměří)", kind: "number", min: 0, step: 0.1, unit: "min", note: "Doporučení vnitřních metodik policie, nikoli metodiky MMR ani zákonný standard." },
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
const VOLUNTEER_FIRE_STATION_HALF_AREA_CODES = new Set([
  "MODEL_BASE_AREA",
  "MODEL_AREA_PER_UNIT",
  "MODEL_SITE_PER_UNIT",
  "MODEL_PAVED_PER_BASE",
  "MODEL_PAVED_PER_UNIT"
]);
const VOLUNTEER_FIRE_STATION_INPUT_CODES = {
  HZS_COVERED_POPULATION: "JSDH_COVERED_POPULATION",
  HZS_AVAILABLE_VEHICLES: "JSDH_AVAILABLE_VEHICLES",
  HZS_COVERAGE_SOURCE: "JSDH_COVERAGE_SOURCE"
};
TARGETED_PLAN_INPUTS[VOLUNTEER_FIRE_STATION_ID] = [
  ...TARGETED_PLAN_INPUTS.VYB_16A
    .filter((field) => field.code !== "HZS_RESPONSE_MINUTES")
    .map((field) => ({
      ...field,
      code: VOLUNTEER_FIRE_STATION_INPUT_CODES[field.code] || field.code,
      label: field.code === "HZS_COVERED_POPULATION"
        ? "Obyvatelé spádu pokrytí stávající jednotkou JSDH"
        : field.code === "HZS_AVAILABLE_VEHICLES"
          ? "Ověřený počet dostupných vozidel / družstev JSDH"
          : field.code === "HZS_COVERAGE_SOURCE"
            ? "Zdroj / datum ověření pěší dostupnosti a pokrytí JSDH"
            : field.label.replaceAll("stanici", "zbrojnici").replaceAll("stanice", "zbrojnice"),
      unit: field.unit.replaceAll("stanici", "zbrojnici"),
      defaultValue: VOLUNTEER_FIRE_STATION_HALF_AREA_CODES.has(field.code)
        ? field.defaultValue / 2 : field.defaultValue,
      note: VOLUNTEER_FIRE_STATION_HALF_AREA_CODES.has(field.code)
        ? "Výchozí scénář používá polovinu výměry HZS; předpoklad lze upravit podle konkrétní zbrojnice."
        : field.note
    })),
  { code: "JSDH_WALK_DISTANCE", label: "Ověřená skutečná pěší vzdálenost k existující zbrojnici (limit 1 000 m)", kind: "number", min: 0, step: 1, unit: "m" }
];
const TARGETED_PLAN_INPUT_CODES = new Set(
  Object.values(TARGETED_PLAN_INPUTS).flat().map((field) => field.code)
);
const TARGETED_FACILITY_ASSUMPTIONS = {
  VYB_01_J: { unitLabel: "míst", rateSource: "RTS 2026, JKSO 801.3 – budovy pro výuku a výchovu" },
  VYB_02C_1: { unitLabel: "žákovských míst", rateSource: "RTS 2026, JKSO 801.3 – budovy pro výuku a výchovu" },
  VYB_02C_2: { unitLabel: "žákovských míst", rateSource: "RTS 2026, JKSO 801.4 – budovy pro vědu, kulturu a osvětu (cenový proxy)" },
  VYB_09A_4: { unitLabel: "lékařských úvazků", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_17: { unitLabel: "lékařských úvazků", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_18: { unitLabel: "lékařských úvazků", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_19: { unitLabel: "lékařských úvazků", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_09B_1: { unitLabel: "ambulancí", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_09B_2: { unitLabel: "ambulancí", rateSource: "RTS 2026, JKSO 801.1 – budovy pro zdravotní péči" },
  VYB_09D: { unitLabel: "výjezdových vozů", rateSource: "RTS 2026, JKSO 801.6 – správní/provozní budovy (cenový proxy)" },
  VYB_16A: { unitLabel: "požárních vozidel", rateSource: "RTS 2026, JKSO 801.6 – správní/provozní budovy (cenový proxy)" },
  [VOLUNTEER_FIRE_STATION_ID]: { unitLabel: "požárních vozidel", rateSource: "RTS 2026, JKSO 801.6 – provozní stavby (cenový proxy; poloviční výměra vůči modelu HZS)" },
  VYB_17A: { unitLabel: "policejních míst", rateSource: "RTS 2026, JKSO 801.6 – budovy pro řízení, správu a administrativu" },
  VYB_17B: { unitLabel: "osob", rateSource: "Scénářový odhad mimo RTS; samostatná sazba úkrytu v RTS 2026 není uvedena" }
};
const TARGETED_RESULT_METRIC_LABELS = {
  VYB_01_J: ["Poptávka nepokrytá ověřenou kapacitou", "Navrhovaný počet dětských skupin", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_02C_1: ["Odhadovaná poptávka po místech", "Poptávka po odečtu ověřené volné kapacity", "Počet modelových zařízení", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_02C_2: ["Odhadovaná poptávka po místech", "Poptávka po odečtu ověřené volné kapacity", "Počet modelových zařízení", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_09A_4: ["Demograficky započtená populace z přírůstku záměru", "Indikativní počet lékařských úvazků", "Poptávka po gynekologii po odečtu ověřené kapacity", "Modelový počet ordinací", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_17: ["Demograficky započtená populace z přírůstku záměru", "Požadovaná modelová potřeba", "Poptávka nepokrytá ověřenou kapacitou", "Modelový počet ordinací", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_18: ["Demograficky započtená populace z přírůstku záměru", "Požadovaná modelová potřeba", "Poptávka nepokrytá ověřenou kapacitou", "Modelový počet ordinací", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_19: ["Demograficky započtená populace z přírůstku záměru", "Požadovaná modelová potřeba", "Poptávka nepokrytá ověřenou kapacitou", "Modelový počet ordinací", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_09B_1: ["Demograficky započtená populace z přírůstku záměru", "Požadovaná modelová potřeba", "Poptávka nepokrytá ověřenou kapacitou", "Modelový počet ambulancí", "Modelová potřeba –", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_09B_2: ["Demograficky započtená populace z přírůstku záměru", "Požadovaná modelová potřeba", "Poptávka nepokrytá ověřenou kapacitou", "Modelový počet ambulancí", "Modelová potřeba –", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"],
  VYB_16A: ["Doplňkový počet vozidel k zajištění", "Modelová potřeba nových stanic po kontrole dojezdu", "Modelový počet hasičských stanic", "Orientační užitná plocha", "Orientační celková podlahová plocha", "Orientační zastavěná plocha", "Orientační obestavěný prostor", "Orientační plocha areálu", "Orientační zpevněná plocha", "Orientační investice včetně VRN a rezervy"],
  [VOLUNTEER_FIRE_STATION_ID]: ["Ověřená pěší vzdálenost k existující zbrojnici", "Pěší dostupnost existující zbrojnice", "Doplňkový počet vozidel k zajištění", "Modelová potřeba nových zbrojnic po kontrole pěší dostupnosti", "Modelový počet hasičských zbrojnic JSDH", "Orientační užitná plocha", "Orientační celková podlahová plocha", "Orientační zastavěná plocha", "Orientační obestavěný prostor", "Orientační plocha areálu", "Orientační zpevněná plocha", "Orientační investice včetně VRN a rezervy"],
  VYB_09D: ["Modelová poptávka po výjezdových vozech", "Modelová potřeba nových základen po kontrole dojezdu", "Modelový počet výjezdových základen", "Orientační užitná plocha", "Orientační celková podlahová plocha", "Orientační zastavěná plocha", "Orientační obestavěný prostor", "Orientační plocha areálu", "Orientační zpevněná plocha", "Orientační investice včetně VRN a rezervy"],
  VYB_17A: ["Scénářový počet policistů k zajištění", "Modelová potřeba nových služeben", "Modelový počet policejních služeben", "Orientační užitná plocha", "Orientační celková podlahová plocha", "Orientační zastavěná plocha", "Orientační obestavěný prostor", "Orientační plocha areálu", "Orientační zpevněná plocha", "Orientační investice včetně VRN a rezervy"],
  VYB_17B: ["Poptávka nepokrytá ověřenou kapacitou úkrytů", "Modelový počet úkrytových modulů", "Orientační užitná plocha", "Orientační obestavěný prostor", "Orientační investice včetně VRN a rezervy"]
};
const isTargetedModelInput = (field) => field.code.startsWith("MODEL_")
  || field.code === "SHELTER_PLANNED_DEMAND";
const TARGETED_FIELD_COLORS = {
  MODEL_DEMOGRAPHIC_SHARE: "area",
  MODEL_DEMOGRAPHIC_COEFFICIENT: "area",
  MODEL_ATTENDANCE_SHARE: "area",
  MODEL_POPULATION_PER_FTE: "area",
  MODEL_LEGAL_PATIENTS_PER_DOCTOR: "area",
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
  VYB_19: [{ profession: "Zubní lékař", group: "Skupina 1", limitMinutes: 35 }],
  VYB_20: [{ profession: "Gynekologie / sdružené ambulantní pracoviště", group: "Ověřit podle odbornosti", limitMinutes: null }],
  VYB_09A_4: [{ profession: "Gynekologie a porodnictví", group: "Skupina 1", limitMinutes: 35 }],
  VYB_09B_1: [
    { profession: "Diabetologie", group: "2", limitMinutes: 45 },
    { profession: "Chirurgie", group: "2", limitMinutes: 45 },
    { profession: "Urologie", group: "2", limitMinutes: 45 },
    { profession: "Interní lékařství", group: "2", limitMinutes: 45 },
    { profession: "Otorinolaryngologie (ORL)", group: "2", limitMinutes: 45 },
    { profession: "Oftalmologie", group: "2", limitMinutes: 45 },
    { profession: "Ortopedie", group: "2", limitMinutes: 45 },
    { profession: "Radiologie (UZ/RTG)", group: "2", limitMinutes: 45 },
    { profession: "Rehabilitační a fyzikální medicína / fyzioterapie", group: "2", limitMinutes: 45 }
  ],
  VYB_09B_2: [
    { profession: "Gastroenterologie", group: "3", limitMinutes: 60 },
    { profession: "Kardiologie", group: "3", limitMinutes: 60 },
    { profession: "Pneumologie a ftizeologie", group: "3", limitMinutes: 60 },
    { profession: "Hematologie a transfuzní lékařství", group: "3", limitMinutes: 60 },
    { profession: "Revmatologie", group: "3", limitMinutes: 60 },
    { profession: "Hemodialýza", group: "3", limitMinutes: 60 },
    { profession: "Psychiatrie / psychologie", group: "3", limitMinutes: 60 },
    { profession: "Logopedie", group: "3", limitMinutes: 60 },
    { profession: "Dermatovenerologie", group: "3", limitMinutes: 60 }
  ]
};
const HEALTHCARE_POPULATION_PER_FTE = {
  VYB_17: { value: 1800, label: "VPL (MMR 2020, optimální)" },
  VYB_18: { value: 4500, label: "PLDD (MMR 2020, optimální)" },
  VYB_19: { value: 2100, label: "stomatologie (MMR 2020, optimální)" },
  VYB_09A_4: { value: 14000, label: "gynekologie (MMR 2020, optimální)" }
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
  POLY_DISCOUNT_COEFFICIENT: "surcharge",
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
  socialServiceChecks: {},
  standardComparisons: {},
  healthChecks: {},
  parkingShares: {},
  realizationSelections: {},
  itemEstimateValues: {},
  brnoInputs: {},
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
  realizedExport: document.querySelector("#realized-export"),
  workbookName: document.querySelector("#workbook-name"),
  notice: document.querySelector("#notice"),
  itemTotal: document.querySelector("#item-total"),
  workbookStatus: document.querySelector("#workbook-status"),
  themeToggle: document.querySelector("#theme-toggle"),
  scenarioStatus: document.querySelector("#scenario-status"),
  resetInputs: document.querySelector("#reset-inputs"),
  projectInputGrid: document.querySelector("#project-input-grid"),
  derivedGrid: document.querySelector("#derived-grid"),
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
  reserveCount: document.querySelector("#reserve-count"),
  reserveList: document.querySelector("#reserve-list"),
  brnoCount: document.querySelector("#brno-count"),
  brnoList: document.querySelector("#brno-list"),
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
      if (id === "VYB_01_J") continue;
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
        manualCoverageCheck: false
      });
      const item = items.at(-1);
      if (USER_ADDED_ITEM_IDS.has(id)) {
        item.userAdded = true;
        item.coverage = "Uživatelská položka";
      }
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
      category: "OSTATNÍ / UŽIVATELSKÉ SCÉNÁŘE",
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
      coverage: "Uživatelská položka",
      legalRules: [],
      recommendedRules: [],
      otherRules: [],
      manualCoverageCheck: false,
      userAdded: true
    });
    addVolunteerFireStationItem(items);
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
      if (equipmentIds.length && equipmentIds.every((id) => id === "VYB_01_J")) continue;
      const resolvedEquipmentIds = equipmentIds.filter((id) => id !== "VYB_01_J");
      const { items: matchingItems, unresolvedIds } = resolveRuleItems(resolvedEquipmentIds, ruleId, items);
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
        ? {
          ...record,
          ...(ruleId === "ZEL-MET-004" ? {
            [RULE_HEADERS.limitation]: "V typech A/B/C porovnejte 300 m se skutečnou pěší trasou k okraji parku. Typ D nemá pevný vzdálenostní limit; ověřte parkovou zeleň v každém obytném sídle. Aplikace pro typ D nezobrazuje odhadnou 300m síť."
          } : {}),
          appLinkedIds: matchingIds,
          appMapping: true
        }
        : record;
      for (const item of matchingItems) {
        const group = rulesByEquipment.get(item.key) || [];
        group.push(linkedRecord);
        rulesByEquipment.set(item.key, group);
      }
    }
    for (const item of items) {
      const itemRules = (rulesByEquipment.get(item.key) || []).filter((rule) =>
        !(item.id === "VYB_04B" && rule[RULE_HEADERS.id] === "ZEL-MET-004"));
      for (const rule of itemRules) {
        if (isLegalRule(rule)) item.legalRules.push(rule);
        else if (isRecommendedRule(rule)) item.recommendedRules.push(rule);
        else item.otherRules.push(rule);
      }
      if (item.id === "VYB_09D") {
        item.legalRules.push({
          [RULE_HEADERS.id]: "ZZS-LEG-001",
          [RULE_HEADERS.equipmentId]: item.id,
          [RULE_HEADERS.equipment]: item.name,
          [RULE_HEADERS.area]: "Zdravotnictví",
          [RULE_HEADERS.type]: "Zákonný požadavek na dostupnost",
          [RULE_HEADERS.condition]: "Dostupnost zdravotnické záchranné služby",
          [RULE_HEADERS.description]: "Čas dojezdu výjezdové skupiny nejpozději do 20 minut od převzetí pokynu k výjezdu, s výjimkami stanovenými zákonem.",
          [RULE_HEADERS.formula]: "Ověřený dojezd ≤ 20 minut.",
          [RULE_HEADERS.inputs]: "Ověřený dojezd v minutách; doložené územní pokrytí.",
          [RULE_HEADERS.output]: "Posouzení dostupnosti ZZS",
          [RULE_HEADERS.rounding]: "Bez zaokrouhlení",
          [RULE_HEADERS.source]: "Zákon č. 374/2011 Sb., o zdravotnické záchranné službě",
          [RULE_HEADERS.location]: "§ 5 – dostupnost zdravotnické záchranné služby",
          [RULE_HEADERS.url]: "https://www.e-sbirka.cz/sb/2011/374",
          [RULE_HEADERS.verification]: "Zákonný limit; aplikace pouze provádí plánovací kontrolu zadaných údajů.",
          [RULE_HEADERS.limitation]: "Zákonné výjimky a krajský plán pokrytí musí být posouzeny poskytovatelem ZZS; limit sám o sobě nezakládá povinnost stavět novou základnu."
        });
      }
      if (item.id === "VYB_04B") item.recommendedRules.push(MMR_PUBLIC_SPACE_RULE);
      if (MMR_SPECIALIST_RULES[item.id]) {
        item.recommendedRules.push(MMR_SPECIALIST_RULES[item.id]);
      }
    }
    addSocialAvailabilityCatalogItems(items);

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
    const constantsRows = worksheetRows(workbook.Sheets.Konstanty_a_Ciselniky);
    const healthcarePopulationPerFte = {
      VYB_17: numericOrNull(cellValue(constantsRows[11], 2)),
      VYB_18: numericOrNull(cellValue(constantsRows[12], 2)),
      VYB_19: numericOrNull(cellValue(constantsRows[13], 2))
    };
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
  if (["ZDR-LEG-001", "ZDR-SPEC-001"].includes(id)) return true;
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
    ["review", elements.reviewList],
    ["reserve", elements.reserveList],
    ["brno", elements.brnoList]
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
    schemaVersion: 17,
    primaryValues: state.primaryValues,
    secondaryValues: state.secondaryValues,
    capacityEdits: state.capacityEdits,
    overrides: state.overrides,
    reviewChecks: state.reviewChecks,
    socialServiceChecks: state.socialServiceChecks,
    standardComparisons: state.standardComparisons,
    healthChecks: state.healthChecks,
    parkingShares: state.parkingShares,
    realizationSelections: state.realizationSelections,
    itemEstimateValues: state.itemEstimateValues,
    brnoInputs: state.brnoInputs,
    scenarioChanged: state.scenarioChanged,
    inputsChanged: state.inputsChanged,
    ...getProjectViewState()
  };
}

function applyProjectScenarioData(data) {
  const savedSchemaVersion = data.schemaVersion || 1;
  const savedPrimaryValues = { ...data.primaryValues };
  if ((data.schemaVersion || 1) < 2) delete savedPrimaryValues.O_sidlo;
  if ((data.schemaVersion || 1) < 4 && savedPrimaryValues.O_sidlo == null) {
    savedPrimaryValues.O_sidlo = 1500;
  }
  state.primaryValues = { ...state.primaryValues, ...savedPrimaryValues };
  delete state.primaryValues.SOCIAL_MUNICIPALITY_POPULATION;
  state.secondaryValues = { ...state.secondaryValues, ...data.secondaryValues };
  if (savedSchemaVersion < 10
    && numericOrNull(state.secondaryValues.POLY_DISCOUNT_COEFFICIENT) === 0.6) {
    state.secondaryValues.POLY_DISCOUNT_COEFFICIENT = 0.7;
  }
  state.capacityEdits = data.capacityEdits || {};
  state.overrides = data.overrides || {};
  state.reviewChecks = data.reviewChecks || {};
  state.socialServiceChecks = data.socialServiceChecks
    && typeof data.socialServiceChecks === "object" ? data.socialServiceChecks : {};
  const serviceCoverage = state.socialServiceChecks.coverage;
  if (serviceCoverage && typeof serviceCoverage === "object"
    && serviceCoverage["community-services"]
    && !serviceCoverage["ambulatory-social-care"]) {
    serviceCoverage["ambulatory-social-care"] = serviceCoverage["community-services"];
  }
  state.standardComparisons = data.standardComparisons || {};
  state.healthChecks = data.healthChecks || {};
  state.parkingShares = data.parkingShares || {};
  state.realizationSelections = data.realizationSelections || {};
  state.itemEstimateValues = data.itemEstimateValues || {};
  if (savedSchemaVersion < 16) {
    for (const item of state.items) {
      if (!item.socialDemographic) continue;
      const values = state.itemEstimateValues[item.key];
      const savedCoefficient = numericOrNull(values?.SOCIAL_DEMOGRAPHIC_COEFFICIENT);
      if (savedCoefficient !== null) {
        values.SOCIAL_DEMOGRAPHIC_COEFFICIENT = savedSchemaVersion < 15
          ? savedCoefficient * (item.id === "VYB_06B" ? 2
            : item.id === "SOCIAL_DISABILITY_SERVICES" ? 1.2 : 0) / 1000
          : savedCoefficient / 1000;
      }
    }
  }
  if (savedSchemaVersion < 17) {
    const previousDefaults = new Map([
      ["VYB_06A", 0.0001],
      ["SOCIAL_LOW_THRESHOLD_DAY_CENTER", 0.0001],
      ["SOCIAL_AMBULATORY_SOCIAL_CARE", 0.0001],
      ["SOCIAL_NIGHT_SHELTER", 0.00005],
      ["SOCIAL_CRISIS_SERVICES", 0.0001],
      ["SOCIAL_HALFWAY_HOUSE", 0.00005],
      ["VYB_06B", 0.002],
      ["SOCIAL_DISABILITY_SERVICES", 0.0012]
    ]);
    for (const item of state.items) {
      const previousDefault = previousDefaults.get(item.id);
      const values = state.itemEstimateValues[item.key];
      const savedCoefficient = numericOrNull(values?.SOCIAL_DEMOGRAPHIC_COEFFICIENT);
      if (item.socialDemographic && previousDefault !== undefined
        && savedCoefficient !== null && Math.abs(savedCoefficient - previousDefault) < 1e-9) {
        values.SOCIAL_DEMOGRAPHIC_COEFFICIENT = item.socialDemographic.coefficient;
      }
    }
  }
  state.brnoInputs = data.brnoInputs && typeof data.brnoInputs === "object"
    ? data.brnoInputs : {};
  const costSurchargeCodes = new Set([
    "ITEM_COST_SURCHARGE", "MODEL_COST_SURCHARGE", "ALT_COST_SURCHARGE", "COST_SURCHARGE"
  ]);
  for (const values of Object.values(state.itemEstimateValues)) {
    if (!values || typeof values !== "object") continue;
    const storedSurchargeCodes = Object.keys(values).filter((code) =>
      /^ITEM_COST_SURCHARGE_(LEGAL|RECOMMENDED)$/.test(code));
    for (const code of [...costSurchargeCodes, ...storedSurchargeCodes]) {
      if (numericOrNull(values[code]) === 0.23) values[code] = 0.08;
    }
  }
  for (const key of Object.keys(state.brnoInputs)) {
    if (key.endsWith("-cost-surcharge")
      && numericOrNull(state.brnoInputs[key]) === 0.23) state.brnoInputs[key] = 0.08;
  }
  for (const key of REMOVED_BRNO_REQUIREMENT_KEYS) {
    delete state.realizationSelections[key];
    delete state.capacityEdits[key];
  }
  for (const key of Object.keys(state.brnoInputs)) {
    if (/^(cars|bikes)-area-|^(cars|bikes)-dwelling-units$/.test(key)
      || ["setback-distance", "setback-no-exception",
        "stormwater-area", "stormwater-custom-limit"].includes(key)) {
      delete state.brnoInputs[key];
    }
  }
  if (savedSchemaVersion < 11) {
    const legalParking = state.items.find((item) => item.id === RESIDENT_PARKING_ID);
    const oldRecommendedKey = legalParking
      ? `${legalParking.key}::${RESIDENT_PARKING_LAW_ID}` : null;
    const recommendedParking = state.items.find((item) =>
      item.id === RESIDENT_PARKING_RECOMMENDED_ID);
    const reserveParking = state.items.find((item) => item.id === RESIDENT_PARKING_RESERVE_ID);
    if (oldRecommendedKey) {
      delete state.overrides[oldRecommendedKey];
      delete state.capacityEdits[oldRecommendedKey];
      delete state.standardComparisons[oldRecommendedKey];
    }
    if (Number(state.realizationSelections[oldRecommendedKey]) === 1) {
      state.realizationSelections[oldRecommendedKey] = 0;
      if (legalParking) state.realizationSelections[legalParking.key] = 0;
      if (recommendedParking) state.realizationSelections[recommendedParking.key] = 1;
    }
    const hasParkingBase = [legalParking, recommendedParking]
      .some((item) => item && Number(state.realizationSelections[item.key]) === 1);
    if (!hasParkingBase && Number(state.realizationSelections[reserveParking?.key]) === 1
      && legalParking) {
      state.realizationSelections[legalParking.key] = 1;
    }
  }
  if (savedSchemaVersion < 12) {
    const legalParking = state.items.find((item) => item.id === RESIDENT_PARKING_ID);
    const oldLegalAlternativeKey = legalParking
      ? `${legalParking.key}::${RESIDENT_PARKING_LAW_ID}` : null;
    const recommendedParking = state.items.find((item) =>
      item.id === RESIDENT_PARKING_RECOMMENDED_ID);
    if (Number(state.realizationSelections[oldLegalAlternativeKey]) === 1 && legalParking) {
      state.realizationSelections[legalParking.key] = 1;
    }
    if (oldLegalAlternativeKey) {
      state.realizationSelections[oldLegalAlternativeKey] = 0;
      delete state.overrides[oldLegalAlternativeKey];
    }
    if (legalParking && recommendedParking
      && Number(state.realizationSelections[legalParking.key]) === 1
      && Number(state.realizationSelections[recommendedParking.key]) === 1) {
      state.realizationSelections[recommendedParking.key] = 0;
    }
  }
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
    && !REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(
      state.items.find((item) => item.key === data.selectedItemKey)?.id)
    ? data.selectedItemKey : "";
  state.items.filter((item) => REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(item.id))
    .forEach((item) => {
      delete state.realizationSelections[item.key];
      delete state.capacityEdits[item.key];
      delete state.overrides[item.key];
      delete state.itemEstimateValues[item.key];
    });
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

function getSelectedStandardMode() {
  const value = getPrimary("REZIM_STD");
  return value === "MAX_STANDARD" || value === "STANDARD" ? value : "ZAKON";
}

function isCompleteStandardMode() {
  return getSelectedStandardMode() !== "ZAKON";
}

function isHighestStandardMode() {
  return getSelectedStandardMode() === "MAX_STANDARD";
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
    O_prah: projectedAreaPopulation !== null ? projectedAreaPopulation + (catchment ?? 0) : null,
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
      const value = field.kind === "derived" ? derivedValues()[field.code]
        : definition.secondary ? getSecondary(field.code) : getPrimary(field.code);
      fields.append(createInputControl(field, value, Boolean(definition.secondary)));
    }
    group.append(fields);
    elements.projectInputGrid.append(group);
  }
  renderDerivedInputs();
  const derivedDetails = document.querySelector(".derived-details");
  if (derivedDetails) derivedDetails.open = state.derivedDetailsOpen;
  elements.resetInputs.disabled = false;
  updateScenarioStatus();
}

function renderDerivedInputs() {
  const values = derivedValues();
  elements.derivedGrid.replaceChildren();
  const editableInputs = createElement("div", "derived-editable-inputs");
  editableInputs.append(createElement("h3", "secondary-heading", "Vstupy pro odhad nových bytů"));
  const editableFields = createElement("div", "derived-editable-fields");
  for (const field of DERIVED_PARAMETER_INPUTS) {
    editableFields.append(createInputControl(field, getPrimary(field.code)));
  }
  editableInputs.append(editableFields);
  elements.derivedGrid.append(editableInputs);
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
  const viewport = captureRequirementViewport(control);
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

  const cards = [...document.querySelectorAll(".requirement-card")]
    .filter((candidate) => candidate.dataset.itemKey === item.key);
  for (const card of cards) {
    const kind = card.dataset.kind;
    if (!kind) continue;
    const estimate = getTargetedPlanEstimate(item, kind === "recommended");
    card.querySelectorAll('[data-input-scope="itemEstimate"][data-input-code]')
      .forEach((otherControl) => {
        if (otherControl.dataset.inputCode === control.dataset.inputCode && otherControl !== control) {
          otherControl.value = control.value;
        }
      });
    card.querySelector(".item-results")?.replaceWith(renderItemMetrics(item, kind));
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
    const quantity = card.querySelector(".requirement-quantity");
    if (quantity) quantity.textContent = formatRequirementSummaryQuantity(
      item, kind, getTargetedPlanSummaryText(item, estimate));
    const cost = estimate.metrics.find((metric) => metric.label.includes("investice"))?.value
      ?? estimate.cost;
    const investment = card.querySelector(".requirement-investment");
    if (investment) investment.textContent = cost === null || cost === undefined
      ? "bez samostatného ocenění" : formatCost(cost);
    const result = card.querySelector(".capacity-result");
    if (result) updateCapacityResultText(result, item, adjustedResult(item, kind), kind);
    if (isHealthcareDemandItem(item)) {
      const demandLabel = item.id === "VYB_09A_4"
        ? "Poptávka po gynekologii po odečtu ověřené kapacity"
        : "Poptávka nepokrytá ověřenou kapacitou";
      card.classList.toggle("is-zero", estimate.metrics.find((metric) =>
        metric.label.startsWith(demandLabel))?.value === 0);
    }
  }
  updateScenarioStatus();
  restoreRequirementViewport(viewport, false);
  return true;
}

function updateSpecialBuildingParkingInput(event) {
  const control = event.target.closest("[data-special-parking-input]");
  if (!control) return;
  const viewport = captureRequirementViewport(control);
  const displayedValue = Number(control.value);
  const min = Number(control.min);
  const max = control.max === "" ? Number.POSITIVE_INFINITY : Number(control.max);
  if (!Number.isFinite(displayedValue) || displayedValue < min || displayedValue > max) {
    control.setCustomValidity("Zadejte hodnotu v povoleném rozsahu.");
    control.reportValidity();
    return;
  }
  control.setCustomValidity("");
  const scale = numericOrNull(control.dataset.specialParkingScale) ?? 1;
  const values = state.itemEstimateValues[SPECIAL_BUILDING_PARKING_KEY] || {};
  values[control.dataset.specialParkingInput] = displayedValue / scale;
  state.itemEstimateValues[SPECIAL_BUILDING_PARKING_KEY] = values;
  const disclosureState = captureRequirementDisclosureState();
  saveScenario(false);
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  restoreRequirementViewport(viewport);
}

function captureRequirementDisclosureState() {
  return [...document.querySelectorAll(".requirement-card")].map((card) => ({
    kind: card.dataset.kind,
    itemKey: card.dataset.itemKey,
    open: card.open,
    sections: [...card.querySelectorAll("details")].map((details) => details.open)
  }));
}

function captureRequirementViewport(target) {
  const card = target?.closest(".requirement-card");
  const list = card?.closest(".requirement-list");
  const active = document.activeElement;
  const descriptor = active?.dataset?.capacityField
    ? {
      recordId: active.dataset.recordId,
      capacityField: active.dataset.capacityField
    }
    : active?.dataset?.capacitySatisfactionItemKey
      ? { satisfactionKey: active.dataset.capacitySatisfactionItemKey }
      : active?.dataset?.specialParkingInput
        ? { specialParkingInput: active.dataset.specialParkingInput }
      : active?.dataset?.inputCode && active?.dataset?.itemKey
        ? { inputCode: active.dataset.inputCode, itemKey: active.dataset.itemKey }
      : null;
  return {
    scrollY: window.scrollY,
    listId: list?.id,
    listScrollTop: list?.scrollTop ?? null,
    descriptor
  };
}

function restoreRequirementViewport(viewport, restoreFocus = true) {
  requestAnimationFrame(() => {
    const list = viewport.listId ? document.getElementById(viewport.listId) : null;
    if (list && viewport.listScrollTop !== null) list.scrollTop = viewport.listScrollTop;
    window.scrollTo(0, viewport.scrollY);
    if (!restoreFocus || !viewport.descriptor) return;
    const control = viewport.descriptor.satisfactionKey
      ? [...document.querySelectorAll("[data-capacity-satisfaction-item-key]")]
        .find((candidate) => candidate.dataset.capacitySatisfactionItemKey
          === viewport.descriptor.satisfactionKey)
      : viewport.descriptor.specialParkingInput
        ? [...document.querySelectorAll("[data-special-parking-input]")]
          .find((candidate) => candidate.dataset.specialParkingInput
            === viewport.descriptor.specialParkingInput)
        : viewport.descriptor.inputCode
          ? [...document.querySelectorAll("[data-input-code][data-item-key]")]
            .find((candidate) => candidate.dataset.itemKey === viewport.descriptor.itemKey
              && candidate.dataset.inputCode === viewport.descriptor.inputCode)
          : [...document.querySelectorAll("[data-record-id][data-capacity-field]")]
            .find((candidate) => candidate.dataset.recordId === viewport.descriptor.recordId
              && candidate.dataset.capacityField === viewport.descriptor.capacityField);
    control?.focus({ preventScroll: true });
  });
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
    if (HIDDEN_DUPLICATE_ITEM_IDS.has(item.id)
      || REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(item.id)) return false;
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
  const catalogItems = state.items.filter((item) => !HIDDEN_DUPLICATE_ITEM_IDS.has(item.id)
    && !REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(item.id));
  const realizationPlan = getRealizationPlan();
  const realizationSummary = createElement("div", "realization-summary");
  realizationSummary.append(createElement("strong", "",
    `Přímo vybráno: ${realizationPlan.directKeys.size}`));
  realizationSummary.append(createElement("span", "",
    "Žádná položka ani infrastruktura se nezařazuje automaticky."));
  realizationSummary.append(createElement("span", "realization-legend",
    "Barevný tvar označuje skupinu; plné zvýraznění znamená výběr k realizaci."));
  const selectableItems = catalogItems.filter((item) =>
    item.id !== RESIDENT_PARKING_LAW_ID
      && item.id !== RESIDENT_PARKING_RECOMMENDED_ID
      && item.id !== RESIDENT_PARKING_RESERVE_ID
      && !isInfrastructureReserveItem(item));
  const allSelected = selectableItems.length > 0
    && selectableItems.every((item) => Number(state.realizationSelections[item.key]) === 1)
    && Number(state.realizationSelections[SPECIAL_BUILDING_PARKING_REALIZATION_KEY]) === 1;
  const realizationToggle = createElement("button", "realization-toggle-all",
    allSelected ? "Zrušit výběr všech položek k realizaci" : "Označit všechny položky k realizaci");
  realizationToggle.type = "button";
  realizationToggle.dataset.realizationToggleAll = "";
  realizationSummary.append(realizationToggle);
  elements.categoryList.append(realizationSummary);
  const allButton = createElement("button", `category-button${!state.selectedCategory && !state.selectedItemKey ? " is-active" : ""}`);
  allButton.type = "button";
  allButton.dataset.category = "";
  allButton.append(createElement("span", "category-swatch all-swatch"));
  allButton.append(createElement("span", "category-name", "Všechny skupiny"));
  allButton.append(createElement("span", "category-count", catalogItems.length));
  elements.categoryList.append(allButton);
  state.categories.forEach((category, index) => {
    const categoryItems = catalogItems.filter((item) => item.category === category);
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
      const standardInactive = isMmrRequirementInactive(item);
      const visuallyInactive = standardInactive;
      const capacitySatisfied = state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true;
      const reserveUnavailable = item.id === RESIDENT_PARKING_RESERVE_ID
        && !getSelectedResidentParkingBase();
      const shape = categoryShape(category);
      const selectionLabel = createElement("label",
        `realization-select-label category-${categoryIndex(category)} shape-${shape}`);
      const selection = document.createElement("input");
      selection.type = "checkbox";
      selection.disabled = capacitySatisfied || reserveUnavailable;
      selection.checked = selected;
      selection.value = selected ? "1" : "0";
      selection.dataset.realizationItemKey = item.key;
      selection.setAttribute("aria-label", `${selected ? "Zrušit výběr" : "Vybrat"}: ${displayCatalogItemName(item)}`);
      selection.title = standardInactive
        ? getInactiveRecommendationText(item)
        : capacitySatisfied
          ? "Uživatel potvrdil dostatečnou kapacitu v okolí."
          : reserveUnavailable
            ? "Nejprve vyberte zákonnou nebo doporučenou variantu parkování."
          : "Označit tuto položku (hodnota 1), nebo její výběr zrušit (hodnota 0).";
      selectionLabel.append(selection);
      const itemButton = createElement("button",
        `category-equipment-button${state.selectedItemKey === item.key ? " is-active" : ""}${selected ? " is-realization-selected" : ""}${capacitySatisfied ? " is-capacity-satisfied" : ""}${visuallyInactive || reserveUnavailable ? " is-standard-inactive" : ""}`,
        displayCatalogItemName(item));
      itemButton.type = "button";
      itemButton.dataset.itemKey = item.key;
      if (standardInactive) itemButton.title = getInactiveRecommendationText(item);
      if (reserveUnavailable) itemButton.title = "Rezervu lze zvolit až po výběru jedné základní varianty parkování.";
      const row = createElement("div",
        `category-equipment-row${capacitySatisfied ? " is-capacity-satisfied" : ""}${visuallyInactive || reserveUnavailable ? " is-standard-inactive" : ""}`);
      row.append(selectionLabel, itemButton);
      items.append(row);
    });
    group.append(items);
    elements.categoryList.append(group);
  });
}

function getRealizationPlan() {
  const parkingReserve = state.items.find((item) => item.id === RESIDENT_PARKING_RESERVE_ID);
  const parkingReserveSelected = parkingReserve
    && Number(state.realizationSelections[parkingReserve.key]) === 1;
  const directItems = state.items.filter((item) =>
    !HIDDEN_DUPLICATE_ITEM_IDS.has(item.id)
    && !REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(item.id)
      &&
    Number(state.realizationSelections[item.key]) === 1
      && state.capacityEdits[item.key]?.satisfiesNearbyCapacity !== true
      && !(INFRASTRUCTURE_ITEM_IDS.has(item.id)
        && Number(state.realizationSelections[getInfrastructureReserveItem(item)?.key]) === 1)
      && !(isInfrastructureReserveItem(item)
        && Number(state.realizationSelections[item.infrastructureReserveSourceKey]) === 1)
      && !(parkingReserveSelected && isResidentParkingBaseItem(item))
      && (!isResidentParkingBaseItem(item)
        || item.key === getSelectedResidentParkingBaseItem()?.key)
      && (item.id !== RESIDENT_PARKING_RESERVE_ID || getSelectedResidentParkingBaseItem()));
  const directKeys = new Set(directItems.map((item) => item.key));
  for (const key of BRNO_REQUIREMENT_KEYS) {
    if (Number(state.realizationSelections[key]) === 1
      && state.capacityEdits[key]?.satisfiesNearbyCapacity !== true) directKeys.add(key);
  }
  if (Number(state.realizationSelections[SPECIAL_BUILDING_PARKING_REALIZATION_KEY]) === 1
    && state.capacityEdits[SPECIAL_BUILDING_PARKING_REALIZATION_KEY]?.satisfiesNearbyCapacity !== true) {
    directKeys.add(SPECIAL_BUILDING_PARKING_REALIZATION_KEY);
  }
  return {
    directKeys,
    selectedKeys: new Set(directKeys)
  };
}

function isResidentParkingBaseItem(item) {
  return [RESIDENT_PARKING_ID, RESIDENT_PARKING_LAW_ID,
    RESIDENT_PARKING_RECOMMENDED_ID].includes(item.id);
}

function getSelectedResidentParkingBaseItem() {
  return [RESIDENT_PARKING_RECOMMENDED_ID, RESIDENT_PARKING_LAW_ID,
    RESIDENT_PARKING_ID].map((id) => state.items.find((item) => item.id === id))
    .find((item) => item && Number(state.realizationSelections[item.key]) === 1) || null;
}

function isInfrastructureReserveItem(item) {
  return Boolean(item.infrastructureReserveSourceKey);
}

function getInfrastructureReserveSource(item) {
  return isInfrastructureReserveItem(item)
    ? state.items.find((candidate) => candidate.key === item.infrastructureReserveSourceKey) || null
    : null;
}

function getInfrastructureReserveItem(item) {
  if (!INFRASTRUCTURE_ITEM_IDS.has(item.id)) return null;
  return state.items.find((candidate) =>
    candidate.infrastructureReserveSourceKey === item.key) || null;
}

function isInfrastructureItem(item) {
  return INFRASTRUCTURE_ITEM_IDS.has(item.id)
    || isInfrastructureReserveItem(item);
}

function getSelectedResidentParkingBase() {
  const source = state.items.find((item) => item.id === RESIDENT_PARKING_ID);
  if (!source) return null;
  const legal = getResidentParkingLegalEstimate(source);
  const recommended = getResidentParkingRecommendedEstimate(source);
  const selected = getSelectedResidentParkingBaseItem();
  if (!selected) return null;
  const quantity = selected.id === RESIDENT_PARKING_RECOMMENDED_ID
    ? recommended?.recommendedStalls : legal?.requiredStalls;
  return Number.isFinite(quantity) ? { item: selected, quantity } : null;
}

function syncRealizationSelectionControls(item) {
  const selected = isResidentParkingBaseItem(item) || item.id === RESIDENT_PARKING_RESERVE_ID
    ? getRealizationPlan().selectedKeys.has(item.key)
    : Number(state.realizationSelections[item.key]) === 1;
  const capacitySatisfied = state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true;
  const standardInactive = isMmrRequirementInactive(item);
  const reserveUnavailable = item.id === RESIDENT_PARKING_RESERVE_ID
    && !getSelectedResidentParkingBase();
  const selectionUnavailable = capacitySatisfied || reserveUnavailable;
  for (const control of document.querySelectorAll(
    `[data-realization-item-key="${CSS.escape(item.key)}"]`
  )) {
    if (control instanceof HTMLInputElement) {
      control.checked = selected;
      control.value = selected ? "1" : "0";
      control.disabled = selectionUnavailable;
      control.setAttribute("aria-label",
        `${selected ? "Zrušit výběr" : "Vybrat"}: ${displayCatalogItemName(item)}`);
      continue;
    }
    control.disabled = selectionUnavailable;
    control.classList.toggle("is-selected", selected);
    control.setAttribute("aria-pressed", String(selected));
    control.setAttribute("aria-label",
      `${selected ? "Zrušit výběr" : "Vybrat"} k realizaci: ${displayCatalogItemName(item)}`);
    control.title = standardInactive
      ? item.socialAvailabilityStandard
        ? "Doporučený standard se na tento typ území nebo populační práh nevztahuje; položku lze přesto zvolit jako volitelnou realizaci."
        : getInactiveRecommendationText(item)
      : capacitySatisfied
        ? "Kapacita okolí je označena jako dostatečná."
        : reserveUnavailable
          ? "Nejprve vyberte zákonnou nebo doporučenou variantu parkování."
      : selected ? "Odebrat z realizace" : "Vybrat k realizaci";
  }
  document.querySelectorAll(".category-equipment-row").forEach((row) => {
    const button = row.querySelector(".category-equipment-button");
    if (button?.dataset.itemKey !== item.key) return;
    button.classList.toggle("is-realization-selected", selected);
    button.classList.toggle("is-capacity-satisfied", capacitySatisfied);
    button.classList.toggle("is-standard-inactive",
      !item.socialAvailabilityStandard && standardInactive || reserveUnavailable);
    row.classList.toggle("is-capacity-satisfied", capacitySatisfied);
    row.classList.toggle("is-standard-inactive",
      !item.socialAvailabilityStandard && standardInactive || reserveUnavailable);
  });
  document.querySelectorAll(".requirement-card[data-item-key]").forEach((card) => {
    if (card.dataset.itemKey !== item.key) return;
    card.classList.toggle("is-realization-selected", selected);
    card.classList.toggle("is-capacity-satisfied", capacitySatisfied);
    card.classList.toggle("is-standard-inactive",
      card.dataset.kind === "recommended" && standardInactive);
    card.classList.toggle("is-parking-reserve-unavailable", reserveUnavailable);
  });
}

function syncParkingReserveControls() {
  for (const item of state.items.filter((candidate) =>
    isResidentParkingBaseItem(candidate) || candidate.id === RESIDENT_PARKING_RESERVE_ID)) {
    syncRealizationSelectionControls(item);
  }
}

function toggleAllRealizations() {
  const selectableItems = state.items.filter((item) => !HIDDEN_DUPLICATE_ITEM_IDS.has(item.id)
    && item.id !== RESIDENT_PARKING_LAW_ID
    && item.id !== RESIDENT_PARKING_RECOMMENDED_ID
    && item.id !== RESIDENT_PARKING_RESERVE_ID
    && !isInfrastructureReserveItem(item));
  const allSelected = selectableItems.length > 0
    && selectableItems.every((item) => Number(state.realizationSelections[item.key]) === 1)
    && Number(state.realizationSelections[SPECIAL_BUILDING_PARKING_REALIZATION_KEY]) === 1;
  const selectAll = !allSelected;
  const disclosureState = captureRequirementDisclosureState();
  const scrollTop = elements.categoryList.scrollTop;
  for (const item of selectableItems) state.realizationSelections[item.key] = selectAll ? 1 : 0;
  state.items.filter(isInfrastructureReserveItem)
    .forEach((item) => { state.realizationSelections[item.key] = 0; });
  for (const item of state.items.filter((candidate) =>
    candidate.id === RESIDENT_PARKING_LAW_ID
      || candidate.id === RESIDENT_PARKING_RECOMMENDED_ID
      || candidate.id === RESIDENT_PARKING_RESERVE_ID)) {
    state.realizationSelections[item.key] = 0;
  }
  state.realizationSelections[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] = selectAll ? 1 : 0;
  renderCategories();
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  elements.categoryList.scrollTop = scrollTop;
  elements.categoryList.querySelector("[data-realization-toggle-all]")?.focus({ preventScroll: true });
  saveScenario(false);
  updateScenarioStatus();
}

function updateRealizationSelection(event) {
  const input = event.target.closest("[data-realization-item-key]");
  if (!input) return;
  const item = state.items.find((candidate) => candidate.key === input.dataset.realizationItemKey);
  if (!item || input.disabled) return;
  setRealizationSelection(item, input.checked);
}

function setSpecialParkingRealizationSelection(selected) {
  state.realizationSelections[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] = selected ? 1 : 0;
  const count = elements.categoryList.querySelector(".realization-summary strong");
  if (count) count.textContent = `Přímo vybráno: ${getRealizationPlan().directKeys.size}`;
  const disclosureState = captureRequirementDisclosureState();
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  saveScenario();
  updateScenarioStatus();
}

function setBrnoRequirementSelection(key, selected) {
  if (!BRNO_REQUIREMENT_KEYS.has(key)) return;
  state.realizationSelections[key] = selected ? 1 : 0;
  const disclosureState = captureRequirementDisclosureState();
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  const count = elements.categoryList.querySelector(".realization-summary strong");
  if (count) count.textContent = `Přímo vybráno: ${getRealizationPlan().directKeys.size}`;
  saveScenario();
  updateScenarioStatus();
}

function setRealizationSelection(item, selected) {
  if (selected && item.id === RESIDENT_PARKING_RESERVE_ID
    && !getSelectedResidentParkingBase()) {
    showNotice("Nejprve vyberte zákonnou nebo doporučenou variantu parkování.", true);
    return;
  }
  const changedItems = applyRealizationSelectionState(item, selected);
  changedItems.forEach((changedItem) => syncRealizationSelectionControls(changedItem));
  syncParkingReserveControls();
  const count = elements.categoryList.querySelector(".realization-summary strong");
  if (count) {
    count.textContent = `Přímo vybráno: ${getRealizationPlan().directKeys.size}`;
  }
  if (state.hideUnrealizedItems) updateRequirementVisibility();
  const disclosureState = captureRequirementDisclosureState();
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  saveScenario();
  updateScenarioStatus();
}

function applyRealizationSelectionState(item, selected) {
  if (selected && item.id === RESIDENT_PARKING_RESERVE_ID
    && !getSelectedResidentParkingBase()) return [];
  const changedItems = [item];
  const infrastructureSource = getInfrastructureReserveSource(item);
  const infrastructureReserve = INFRASTRUCTURE_ITEM_IDS.has(item.id)
    ? getInfrastructureReserveItem(item) : null;
  const oppositeInfrastructureItem = infrastructureSource || infrastructureReserve;
  if (selected && oppositeInfrastructureItem
    && Number(state.realizationSelections[oppositeInfrastructureItem.key]) === 1) {
    state.realizationSelections[oppositeInfrastructureItem.key] = 0;
    changedItems.push(oppositeInfrastructureItem);
  }
  if (selected && isResidentParkingBaseItem(item)) {
    for (const other of state.items) {
      if (other.key !== item.key && isResidentParkingBaseItem(other)
        && Number(state.realizationSelections[other.key]) === 1) {
        state.realizationSelections[other.key] = 0;
        changedItems.push(other);
      }
    }
    const reserveItem = state.items.find((candidate) =>
      candidate.id === RESIDENT_PARKING_RESERVE_ID);
    if (reserveItem && Number(state.realizationSelections[reserveItem.key]) === 1) {
      state.realizationSelections[reserveItem.key] = 0;
      changedItems.push(reserveItem);
    }
  }
  state.realizationSelections[item.key] = selected ? 1 : 0;
  if (!selected && isResidentParkingBaseItem(item) && !getSelectedResidentParkingBaseItem()) {
    const reserveItem = state.items.find((candidate) =>
      candidate.id === RESIDENT_PARKING_RESERVE_ID);
    if (reserveItem && Number(state.realizationSelections[reserveItem.key]) === 1) {
      state.realizationSelections[reserveItem.key] = 0;
      changedItems.push(reserveItem);
    }
  }
  return changedItems;
}

function toggleRealizationScope(scope) {
  const controls = [...scope.querySelectorAll("button[data-realization-item-key]")]
    .filter((control) => !control.disabled);
  const uniqueControls = [...new Map(controls.map((control) =>
    [control.dataset.realizationItemKey, control])).values()];
  if (!uniqueControls.length) return;
  const selected = !uniqueControls.every((control) =>
    control.getAttribute("aria-pressed") === "true");
  const disclosureState = captureRequirementDisclosureState();
  const changedItems = new Map();
  for (const control of uniqueControls) {
    const key = control.dataset.realizationItemKey;
    if (key === SPECIAL_BUILDING_PARKING_REALIZATION_KEY) {
      state.realizationSelections[key] = selected ? 1 : 0;
      continue;
    }
    if (BRNO_REQUIREMENT_KEYS.has(key)) {
      state.realizationSelections[key] = selected ? 1 : 0;
      continue;
    }
    const item = state.items.find((candidate) => candidate.key === key);
    if (!item) continue;
    for (const changedItem of applyRealizationSelectionState(item, selected)) {
      changedItems.set(changedItem.key, changedItem);
    }
  }
  changedItems.forEach((item) => syncRealizationSelectionControls(item));
  syncParkingReserveControls();
  const count = elements.categoryList.querySelector(".realization-summary strong");
  if (count) count.textContent = `Přímo vybráno: ${getRealizationPlan().directKeys.size}`;
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  saveScenario();
  updateScenarioStatus();
}

function hasActiveUserOverride(override) {
  const quantity = numericOrNull(override?.quantity);
  return quantity !== null && quantity >= 0;
}

function recordUserOverrideChange(item, kind) {
  const override = state.overrides[item.key];
  if (!hasActiveUserOverride(override)) return;
  if (!Number.isFinite(override.baselineQuantity)) {
    override.baselineQuantity = getPlannedQuantity(item, kind, false);
  }
  const history = Array.isArray(override.history) ? override.history : [];
  const previous = history.length
    ? history[history.length - 1].quantity
    : override.baselineQuantity;
  const entry = {
    previousQuantity: previous,
    quantity: override.quantity,
    author: override.author?.trim() || "Neuvedeno",
    reason: override.note?.trim() || "Neuveden",
    changedAt: new Date().toISOString()
  };
  history.push(entry);
  override.history = history;
  override.updatedAt = entry.changedAt;
}

function getUserOverrideAuditSummary(override, unit) {
  const history = Array.isArray(override?.history) ? override.history : [];
  const latest = history[history.length - 1];
  if (!latest) {
    return hasActiveUserOverride(override)
      ? "Uživatelský návrh je aktivní; autor a poznámka jsou nepovinné."
      : "Zákonný či metodický výpočet zůstává výchozí hodnotou. Zadejte pouze uživatelské návrhové množství.";
  }
  const timestamp = new Date(latest.changedAt).toLocaleString("cs-CZ");
  const auditContext = [
    latest.author && latest.author !== "Neuvedeno" ? latest.author : "",
    latest.reason && latest.reason !== "Neuveden" ? latest.reason : ""
  ].filter(Boolean).join("; ");
  return `Výchozí model: ${formatNumber(override.baselineQuantity)} ${text(unit, "jednotek")}. Poslední změna ${timestamp}${auditContext ? ` (${auditContext})` : ""}: ${formatNumber(latest.previousQuantity)} → ${formatNumber(latest.quantity)} ${text(unit, "jednotek")}.`;
}

function getPlannedQuantity(item, kind = "", applyUserOverride = true) {
  const override = state.overrides[item.key];
  if (applyUserOverride && hasActiveUserOverride(override)) return override.quantity;
  if (kind === "recommended" && MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) {
    return getMmrSchoolRecommendation(item)?.requiredCapacity ?? null;
  }
  if (isResidentParkingItem(item)) {
    const legal = getResidentParkingLegalEstimate(item);
    const recommended = getResidentParkingRecommendedEstimate(item);
    if (item.id === RESIDENT_PARKING_ID) return kind === "legal" ? legal?.requiredStalls ?? null : null;
    if (item.id === RESIDENT_PARKING_LAW_ID) {
      return kind === "recommended" ? legal?.requiredStalls ?? null : null;
    }
    if (item.id === RESIDENT_PARKING_RECOMMENDED_ID) {
      return kind === "recommended" ? recommended?.recommendedStalls ?? null : null;
    }
    if (item.id === RESIDENT_PARKING_RESERVE_ID) {
      return kind === "recommended" ? recommended?.reserveStalls ?? null : null;
    }
  }
  if (kind === "recommended" && isMmrWasteYard(item)) {
    return getSpatialPointRequirement(item, kind) ?? item.modelProposed;
  }
  const standardTarget = getAppliedStandardTarget(item);
  const planned = standardTarget !== null ? standardTarget : item.proposed;
  if (kind !== "recommended" || !hasSpatialPointFloor(item)) return planned;
  const spatialCount = getSpatialPointRequirement(item, kind);
  if (!Number.isFinite(spatialCount)) return planned;
  const capacityPerPoint = getSpatialCapacityPerPoint(item);
  const spatialFloor = capacityPerPoint === null ? null : capacityPerPoint * spatialCount;
  if (!Number.isFinite(spatialFloor)) return planned;
  return planned === null || planned === undefined
    ? spatialFloor : Math.max(planned, spatialFloor);
}

function getSpatialPointRequirement(item, kind = "recommended") {
  if (kind !== "recommended" || !hasSpatialPointFloor(item)) return null;
  const count = calculateSpatialEstimate(item)?.count;
  return Number.isFinite(count) ? Math.max(item.modelProposed ?? 0, count) : null;
}

function getNetworkFacilityCount(item, kind) {
  if (kind !== "recommended") return null;
  if (MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) {
    const recommendation = getMmrSchoolRecommendation(item);
    return recommendation?.spatialCapacityFloor > 0
      && recommendation.spatialPointCount > 1
      ? recommendation.spatialPointCount : null;
  }
  if (!hasSpatialPointFloor(item) || isMmrWasteYard(item)) return null;
  const count = calculateSpatialEstimate(item)?.count;
  return Number.isFinite(count) && count > 1 ? count : null;
}

function formatRequirementSummaryQuantity(item, kind, quantity) {
  const quantityText = String(quantity ?? "")
    .replace(/\s+\([^()]*\szařízení\)$/, "");
  const count = getNetworkFacilityCount(item, kind);
  return count === null || !(getExportNumericValue(quantityText) > 0)
    ? quantityText
    : `${quantityText} (${formatNumber(count)} zařízení)`;
}

function getRequirementSummaryUnit(item) {
  return item.id === "VYB_16" ? "lékárna" : getRequirementUnit(item);
}

function formatRequirementSummaryAmount(item, quantity) {
  const unitForms = {
    VYB_03A: ["hřiště", "hřiště", "hřišť"],
    VYB_03B: ["hřiště", "hřiště", "hřišť"],
    VYB_03C: ["sportoviště", "sportoviště", "sportovišť"],
    VYB_16: ["lékárna", "lékárny", "lékáren"]
  }[item.id];
  const unit = unitForms
    ? czechCountUnit(quantity, ...unitForms)
    : getRequirementSummaryUnit(item);
  return `${formatNumber(quantity, 2)} ${text(unit, "")}`;
}

function getSpatialCapacityPerPoint(item) {
  if (isMmrWasteYard(item)) return 1;
  if (!CAPACITY_QUANTITY_IDS.has(item.id)) return 1;
  return item.modelProposed > 0 && item.demand > 0
    ? item.demand / item.modelProposed : null;
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

function getInfrastructureReservePercent(item) {
  const saved = numericOrNull(state.itemEstimateValues[item.key]?.INFRA_RESERVE_PERCENT);
  return saved === null ? 0.1 : Math.min(1, Math.max(0, saved));
}

function getResidentParkingTerritoryType() {
  return String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase() ?? "A";
}

function isMmrSchoolPresenceRecommended(item, localPopulation) {
  const territory = getResidentParkingTerritoryType();
  if (item.id === "VYB_02C_1") {
    if (!Number.isFinite(localPopulation)) return null;
    return ["A", "B"].includes(territory) && localPopulation > 20000;
  }
  if (item.id === "VYB_02C_2") {
    if (territory === "A") return true;
    if (!Number.isFinite(localPopulation)) return null;
    return ["B", "C"].includes(territory) && localPopulation > 5000;
  }
  return null;
}

function isMmrRequirementInactive(item) {
  const territory = getResidentParkingTerritoryType();
  const localPopulation = numericOrNull(derivedValues().O_uzemi_cil);
  if (item.id === "VYB_17A") return territory === "D";
  const schoolPresenceRecommended = isMmrSchoolPresenceRecommended(item, localPopulation);
  if (schoolPresenceRecommended !== null) return !schoolPresenceRecommended;
  if (item.socialAvailabilityStandard) {
    const standard = item.socialAvailabilityStandard;
    if (!standard.territories.includes(territory)) return true;
    if (standard.minimumPopulation !== undefined) {
      const municipalityPopulation = numericOrNull(derivedValues().O_prah);
      return municipalityPopulation !== null
        && municipalityPopulation <= standard.minimumPopulation;
    }
    return false;
  }
  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
    return localPopulation !== null && getMmrSchoolAccessRule(item).mode === "manual";
  }
  if (MMR_HEALTHCARE_GROUP_1_ITEM_IDS.has(item.id)) {
    return ["A", "B", "C"].includes(territory)
      && localPopulation !== null && localPopulation <= 2000;
  }
  if (["VYB_09B_1", "VYB_09B_2"].includes(item.id)) {
    return !["A", "B"].includes(territory)
      || territory === "B" && localPopulation !== null && localPopulation <= 5000;
  }
  return item.id === "VYB_05B";
}

function getInactiveRecommendationText(item) {
  return item.id === "VYB_17A"
    ? "V typu D se doporučení interních metodik policie neměří; modelovanou potřebu lze přesto zvolit jako vlastní realizaci."
    : "Položka je mimo práh doporučení MMR; modelovanou potřebu lze přesto zvolit jako vlastní realizaci.";
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
  const recommendedStalls = Math.max(legal.requiredStalls, modeledStalls);
  const selectedBase = getSelectedResidentParkingBaseItem();
  const selectedBaseStalls = selectedBase
    ? selectedBase.id === RESIDENT_PARKING_RECOMMENDED_ID
      ? recommendedStalls : legal.requiredStalls
    : null;
  const reservePercent = getResidentParkingReservePercent();
  const reserveStalls = selectedBaseStalls === null
    ? null : Math.ceil(selectedBaseStalls * reservePercent);
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
    recommendedStalls,
    modeledApartmentStalls,
    modeledFamilyHomeStalls,
    baseStalls: selectedBaseStalls,
    selectedBaseStalls,
    reservePercent,
    reserveStalls,
    totalWithReserve: selectedBaseStalls === null || reserveStalls === null
      ? null : selectedBaseStalls + reserveStalls
  };
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
  return isCompleteStandardMode()
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
  if (kind === "recommended" && !isCompleteStandardMode()) {
    return comparison.delta / comparison.base;
  }
  if (isCompleteStandardMode() && hasAppliedStandardTarget(item)) {
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

function getRealizedFacilityMetrics(item, kind = "recommended") {
  const estimate = getProjectPlanningEstimate(item, kind);
  const facilityPlan = estimate?.facilityPlan;
  if (REALIZATION_OUTDOOR_IDS.has(item.id)) {
    const siteArea = numericOrNull(estimate?.siteArea)
      ?? numericOrNull(estimate?.usableArea)
      ?? numericOrNull(item.parcelArea)
      ?? numericOrNull(item.usableArea)
      ?? 0;
    return {
      usableArea: 0, totalFloorArea: 0, volume: 0, footprintArea: 0,
      siteArea,
      pavedArea: numericOrNull(estimate?.pavedArea)
        ?? numericOrNull(item.pavedArea) ?? 0
    };
  }
  if (estimate?.type === "public-space") {
    return {
      usableArea: 0, totalFloorArea: 0, volume: 0, footprintArea: 0,
      siteArea: estimate.minimumArea,
      pavedArea: estimate.pavedArea
    };
  }
  const usableArea = numericOrNull(estimate?.usableArea)
    ?? numericOrNull(facilityPlan?.usableArea)
    ?? numericOrNull(item.usableArea);
  const totalFloorFactor = numericOrNull(estimate?.totalFloorFactor)
    ?? numericOrNull(facilityPlan?.totalFloorFactor)
    ?? numericOrNull(getItemEstimateValue(item, "ITEM_TOTAL_FLOOR_FACTOR"))
    ?? 1;
  const volumePerArea = numericOrNull(estimate?.volumePerArea)
    ?? numericOrNull(facilityPlan?.volumePerFloorArea)
    ?? numericOrNull(getItemEstimateValue(item, "ITEM_VOLUME_PER_AREA"))
    ?? 3.3;
  const sourceVolume = numericOrNull(estimate?.volume)
    ?? numericOrNull(facilityPlan?.volume)
    ?? numericOrNull(item.volume);
  const totalFloorArea = numericOrNull(estimate?.totalFloorArea)
    ?? numericOrNull(facilityPlan?.totalFloorArea)
    ?? (usableArea === null ? sourceVolume === null ? null : sourceVolume / volumePerArea
      : usableArea * totalFloorFactor);
  const volume = sourceVolume ?? (totalFloorArea === null ? null : totalFloorArea * volumePerArea);
  const footprintArea = numericOrNull(estimate?.footprintArea)
    ?? numericOrNull(facilityPlan?.footprintArea)
    ?? (totalFloorArea === null ? null
      : totalFloorArea / (totalFloorFactor > 0 ? totalFloorFactor : 1));
  const siteArea = numericOrNull(estimate?.siteArea)
    ?? numericOrNull(facilityPlan?.siteArea)
    ?? numericOrNull(item.parcelArea);
  const pavedArea = numericOrNull(estimate?.pavedArea)
    ?? numericOrNull(facilityPlan?.pavedArea)
    ?? numericOrNull(item.pavedArea)
    ?? (siteArea === null || footprintArea === null ? 0
      : Math.max(0, siteArea - footprintArea)
        * (numericOrNull(getProjectEstimateValue(item, "K_ZPEVNEN")) ?? 0));
  return { usableArea, totalFloorArea, volume, footprintArea, siteArea, pavedArea };
}

function getAggregateProjectSurfaceMetrics() {
  const realizationPlan = getRealizationPlan();
  const residential = derivedValues();
  const residentialFloorArea = numericOrNull(residential.HPP_byty) ?? 0;
  const residentialRoofArea = numericOrNull(residential.areaWhole) ?? 0;
  const residentialVolume = residentialFloorArea
    * (numericOrNull(getSecondary("BYTY_OBESTAVENI_NA_M2")) ?? 3.9);
  let totalFloorArea = residentialFloorArea;
  let totalFacilityVolume = 0;
  let roofArea = residentialRoofArea;
  let pavedArea = 0;
  for (const candidate of state.items) {
    if (!realizationPlan.selectedKeys.has(candidate.key)
      || TECHNICAL_ALTERNATIVE_IDS.has(candidate.id)
      || isInfrastructureReserveItem(candidate)
      || REALIZATION_PARKING_IDS.has(candidate.id)
      || REALIZATION_NON_BUILDING_IDS.has(candidate.id)
      || candidate.manualCoverageCheck && candidate.id !== "VYB_17B") continue;
    const metrics = getRealizedFacilityMetrics(candidate);
    totalFloorArea += numericOrNull(metrics.totalFloorArea) ?? 0;
    totalFacilityVolume += numericOrNull(metrics.volume) ?? 0;
    roofArea += numericOrNull(metrics.footprintArea) ?? 0;
    pavedArea += numericOrNull(metrics.pavedArea) ?? 0;
  }
  return {
    totalFloorArea,
    totalVolume: residentialVolume + totalFacilityVolume,
    totalFacilityVolume,
    roofArea,
    pavedArea,
    runoffArea: roofArea + pavedArea
  };
}

function getAggregateProjectEnergyMetrics(item, includeRealizedFacilities = true) {
  const realizationPlan = getRealizationPlan();
  let annualElectricity = 0;
  let peakElectricalLoad = 0;
  let modeledNonElectricEnergy = 0;
  let modeledEquipmentCount = 0;
  let fallbackFloorArea = 0;
  let fallbackHeatingArea = 0;
  let fallbackBuildingVolume = 0;
  let projectFloorArea = 0;
  const addEnergySource = (candidate, housingEstimate = null) => {
    const metrics = housingEstimate ? null : getRealizedFacilityMetrics(candidate);
    const usableArea = numericOrNull(housingEstimate?.portfolioArea)
      ?? numericOrNull(metrics?.usableArea)
      ?? numericOrNull(candidate.usableArea);
    const totalFloorArea = numericOrNull(housingEstimate?.portfolioArea)
      ?? numericOrNull(metrics?.totalFloorArea)
      ?? usableArea;
    if (totalFloorArea === null || totalFloorArea <= 0) return;
    const buildingVolume = numericOrNull(metrics?.volume)
      ?? (housingEstimate
        ? totalFloorArea * (numericOrNull(getSecondary("BYTY_OBESTAVENI_NA_M2")) ?? 3.9)
        : totalFloorArea * 3.3);
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
      fallbackBuildingVolume += buildingVolume;
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
    fallbackBuildingVolume += residentialFloorArea
      * (numericOrNull(getSecondary("BYTY_OBESTAVENI_NA_M2")) ?? 3.9);
    fallbackHeatingArea += residentialFloorArea;
  }
  if (includeRealizedFacilities) {
    for (const candidate of state.items) {
      if (!realizationPlan.selectedKeys.has(candidate.key)
        || TECHNICAL_ALTERNATIVE_IDS.has(candidate.id)
        || isInfrastructureReserveItem(candidate)
        || REALIZATION_NON_BUILDING_IDS.has(candidate.id)
        || candidate.manualCoverageCheck && candidate.id !== "VYB_17B") continue;
      addEnergySource(candidate);
    }
  }
  const energyDensity = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_DENSITY_W_M3"))
    ?? (numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_DENSITY_W_M2")) ?? 0) / 4.2;
  const peakToAverageRatio = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO")) ?? 4;
  const fallbackPeakElectricalLoad = fallbackBuildingVolume * energyDensity / 1000;
  const fallbackAnnualElectricity = peakToAverageRatio > 0
    ? fallbackPeakElectricalLoad * 8760 / peakToAverageRatio : 0;
  const fallbackHeatDensity = numericOrNull(getItemEstimateValue(item, "BOILER_HEAT_DENSITY_KWH_M3_Y"))
    ?? (numericOrNull(getItemEstimateValue(item, "BOILER_HEAT_DENSITY_KWH_M2_Y")) ?? 0) / 4.2;
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
    fallbackBuildingVolume,
    fallbackHeatDemand: fallbackBuildingVolume * fallbackHeatDensity,
    energyDensity,
    fallbackHeatDensity,
    fallbackHeatingArea,
    projectFloorArea
  };
}

function getItemEstimateDefaultQuantity(item) {
  if (item.socialDemographic) {
    const population = numericOrNull(derivedValues().O_prah);
    const coefficient = numericOrNull(getItemEstimateValue(item,
      "SOCIAL_DEMOGRAPHIC_COEFFICIENT"));
    if (population === null || coefficient === null) return 0;
    return Math.ceil(population * coefficient);
  }
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
  if (item.socialDemographic) return getItemEstimateDefaultQuantity(item);
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

function normalizeCostSurcharge(value) {
  const surcharge = numericOrNull(value);
  return surcharge === 0.23 ? 0.08 : surcharge;
}

function getItemEstimateDefault(item, code) {
  if (code === "SOCIAL_DEMOGRAPHIC_COEFFICIENT") {
    return item.socialDemographic?.coefficient ?? 0;
  }
  const quantity = getItemEstimateDefaultQuantity(item);
  const scenario = ITEM_ESTIMATE_SCENARIOS[item.id];
  const perUnit = (value) => {
    const number = numericOrNull(value);
    return number !== null && number > 0 && quantity > 0 ? number / quantity : null;
  };
  switch (code) {
    case "ITEM_QUANTITY": return getItemEstimateDefaultQuantity(item);
    case "SUPERMARKET_MIN_POPULATION": return 3000;
    case "SUPERMARKET_POPULATION_PER_STORE": return 3000;
    case "PUBLIC_SPACE_QUALIFIED_AREA_HA": return numericOrNull(getPrimary("S_zastav")) ?? 0;
    case "PUBLIC_SPACE_PAVED_SHARE": return 1;
    case "PUBLIC_SPACE_PAVING_RATE_M2": return 2500;
    case "PARK_GREEN_AREA_HA": return 0.5;
    case "PARK_LANDSCAPING_RATE_M2": return 500;
    case "ITEM_USABLE_AREA_PER_UNIT":
      if (item.socialDemographic) return item.socialDemographic.usableAreaPerUnit;
      if (item.id === "VYB_04B") return 5000;
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
      if (item.socialDemographic) return item.socialDemographic.siteAreaPerUnit;
      if (item.id === "VYB_04B") return 5000;
      if (item.id === "VYB_15B_2") return 1.2;
      if (scenario) return scenario.siteArea;
      return perUnit(item.parcelArea) ?? perUnit(item.areaPerPoint);
    case "ITEM_PAVED_AREA_PER_UNIT":
      if (item.id === "VYB_04B") return 600;
      if (item.id === "VYB_15B_2") return 1.5;
      if (scenario) return scenario.pavedArea;
      return perUnit(item.pavedArea);
    case "ITEM_RTS_RATE": {
      const constructionRate = getRtsConstructionRate(item);
      if (constructionRate !== undefined) return constructionRate;
      if (item.socialDemographic) {
        return numericOrNull(item.rtsPricing?.rate)
          ?? numericOrNull(item.unitPrice)
          ?? estimateComparableUnitRate(item);
      }
      if (item.id === "VYB_04B") return 2500000;
      if (item.id === "VYB_15B_2") return 10100;
      if (item.id === "VYB_10B_1") return 13645;
      if (item.rtsPricing?.rate) return numericOrNull(item.rtsPricing.rate);
      if (numericOrNull(item.unitPrice) > 0) return numericOrNull(item.unitPrice);
      if (scenario) return 13645;
      const sourceQuantity = getItemEstimateDefaultQuantity(item);
      if (numericOrNull(item.cost) > 0 && sourceQuantity > 0) {
        return numericOrNull(item.cost) / sourceQuantity / 1.08;
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
      if (item.id === "VYB_10B_1") return 0.08;
      const sourceSurcharge = normalizeCostSurcharge(item.rtsPricing?.surcharge);
      if (sourceSurcharge !== null) return sourceSurcharge;
      const legacyFactor = numericOrNull(getSecondary("K_CENA_SKOLY"));
      return SCHOOL_CAPACITY_ITEM_IDS.has(item.id) && legacyFactor !== null
        ? normalizeCostSurcharge(Math.max(0, legacyFactor - 1)) : 0.08;
    }
    default: return null;
  }
}

function getAlternativeEstimateDefault(code) {
  switch (code) {
    case "ALT_RAINFALL_DEPTH_MM": return 15;
    case "ALT_RUNOFF_COEFFICIENT": return 0.9;
    case "ALT_ROOF_RUNOFF_COEFFICIENT": return 0.95;
    case "ALT_PAVED_RUNOFF_COEFFICIENT": return 0.65;
    case "ALT_RAINWATER_RATE": return 13600;
    case "ALT_WATER_LITERS_PER_PERSON_DAY": return 87;
    case "ALT_FACILITY_WATER_LITERS_PER_M3_DAY": return 1.5;
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
    case "ALT_ENERGY_DENSITY_W_M3": return 12;
    case "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO": return 4;
    case "BOILER_HEAT_DENSITY_KWH_M2_Y": return 60;
    case "BOILER_HEAT_DENSITY_KWH_M3_Y": return 14.3;
    case "BOILER_HEATING_SHARE": return 1;
    case "BOILER_FULL_LOAD_HOURS": return 1800;
    case "BOILER_DESIGN_RESERVE": return 1.2;
    case "BOILER_UNIT_CAPACITY_KW": return 500;
    case "BOILER_UNIT_COST": return null;
    case "ALT_POWER_CABLE_RATE": return 2569.6;
    case "ALT_TRANSFORMER_CAPACITY_KW": return 2500;
    case "ALT_TRANSFORMER_RATE": return 3000000;
    case "ALT_COST_SURCHARGE": return 0.08;
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
    return state.itemEstimateValues[item.key]?.[code]
  || "Samostatná nová hala požadována";
  }
  const alternativeField = ALTERNATIVE_ESTIMATE_INPUTS.find((candidate) => candidate.code === code);
  if (alternativeField) return getAlternativeEstimateDefault(code);
  if (targetedField) return targetedField.defaultValue ?? fallback;
  const supermarketField = SUPERMARKET_ESTIMATE_INPUTS.find((candidate) => candidate.code === code);
  if (supermarketField) return getItemEstimateDefault(item, code);
  const field = ITEM_ESTIMATE_INPUTS.find((candidate) => candidate.code === code);
  return field ? getItemEstimateDefault(item, code) : fallback;
}

function getRtsBuildingGroup(item) {
  return RTS_801_GROUP_BY_ITEM_ID[item.id];
}

function getRtsConstructionOptions(item) {
  const rates = getRtsConstructionRates(item);
  return rates ? Object.keys(rates).sort((left, right) => Number(left) - Number(right)) : [];
}

function getRtsConstructionRates(item) {
  const group = getRtsBuildingGroup(item);
  return group === "811.5" ? RTS_811_5_CONSTRUCTION_RATES : RTS_801_CONSTRUCTION_RATES[group];
}

function getRtsConstructionRate(item) {
  const buildingGroup = getRtsBuildingGroup(item);
  if (!buildingGroup) return undefined;
  const constructionCode = getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code);
  const rates = getRtsConstructionRates(item);
  return Object.prototype.hasOwnProperty.call(rates, constructionCode)
    ? rates[constructionCode] : null;
}

function getParkingEstimateValue(item, kind, code) {
  const storedCode = `${code}_${kind.toUpperCase()}`;
  const saved = state.itemEstimateValues[item.key]?.[storedCode];
  if (saved !== null && saved !== undefined) return saved;
  const sourceItem = item.parkingSourceKey
    ? state.items.find((candidate) => candidate.key === item.parkingSourceKey) || item
    : item;
  switch (code) {
    case "PARKING_BUILDING_SHARE":
      if (item.id === RESIDENT_PARKING_RESERVE_ID) {
        return getSelectedResidentParkingBase()?.item.id === RESIDENT_PARKING_RECOMMENDED_ID
          ? 0.8 : 0.2;
      }
      return kind === "legal" || item.id === RESIDENT_PARKING_LAW_ID ? 0.2 : 0.8;
    case "PARKING_VOLUME_PER_STALL_M3": return 18;
    case "PARKING_BUILDING_RATE_M3": return getRtsConstructionRate(sourceItem);
    case "PARKING_SURFACE_RATE_M3": return 7500;
    case "PARKING_SURFACE_AREA_PER_STALL_M2": return item.id === "VYB_15A_1"
      ? numericOrNull(getPrimary("PARK_PAVE_M2")) ?? 25 : 25;
    case "ITEM_COST_SURCHARGE": return 0.08;
    default: return null;
  }
}

function getTargetedPlanCostRate(item) {
  if (["VYB_09D", "VYB_16A", VOLUNTEER_FIRE_STATION_ID, "VYB_17A"].includes(item.id)) {
    return numericOrNull(getTargetedPlanInput(item, "MODEL_COST_PER_M3"));
  }
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
    if (code === "MODEL_COST_PER_M3"
      && ["VYB_09D", "VYB_16A", VOLUNTEER_FIRE_STATION_ID, "VYB_17A"].includes(item.id)) {
      return numericOrNull(getRtsConstructionRate(item))
        ?? numericOrNull((TARGETED_PLAN_INPUTS[item.id] || [])
          .find((candidate) => candidate.code === code)?.defaultValue);
    }
    if (code === "MODEL_POPULATION_PER_FTE" && ["VYB_17", "VYB_19"].includes(item.id)) {
      return getHealthcareStaffingBasis(item)?.value
        ?? numericOrNull((TARGETED_PLAN_INPUTS[item.id] || [])
          .find((candidate) => candidate.code === code)?.defaultValue);
    }
    const field = (TARGETED_PLAN_INPUTS[item.id] || [])
      .find((candidate) => candidate.code === code);
    return numericOrNull(field?.defaultValue);
  }
  return numericOrNull(saved);
}

function getDoctorPopulationPerFte(item, recommended) {
  const legalBaseline = getTargetedPlanInput(item, "MODEL_LEGAL_PATIENTS_PER_DOCTOR");
  if (!recommended) return legalBaseline;
  return getTargetedPlanInput(item, "MODEL_POPULATION_PER_FTE") ?? legalBaseline;
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

function getMmrSchoolAccessRule(item) {
  const territory = getResidentParkingTerritoryType();
  const localPopulation = numericOrNull(derivedValues().O_uzemi_cil);
  const standard = MMR_SCHOOL_STANDARDS[item.id];
  if (!standard || localPopulation === null) return { mode: "manual", summary: "Dostupnost ověřte ručně." };

  if (item.id === "VYB_01") {
    if (localPopulation > 1000 && ["A", "B", "C", "D"].includes(territory)) {
      const distance = getPrimary("KOMPAKT") === "1" ? 400 : 600;
      return {
        mode: "distance",
        distance,
        density: standard.densityByDistance[distance],
        summary: `skutečná pěší vzdálenost ${formatNumber(distance)} m${distance === 400 ? " v souvislé kompaktní zástavbě" : ""}`
      };
    }
    if (territory === "D" && localPopulation < 1000) {
      return { mode: "vhd", minutes: 30, summary: "VHD do 30 minut" };
    }
  } else if (item.id === "VYB_02A") {
    if (territory === "A") {
      return { mode: "distance", distance: 600, density: standard.densityByDistance[600], summary: "skutečná pěší vzdálenost 600 m" };
    }
    if (["B", "C", "D"].includes(territory) && localPopulation > 2000) {
      return { mode: "distance", distance: 800, density: standard.densityByDistance[800], summary: "skutečná pěší vzdálenost 800 m" };
    }
    if (["B", "D"].includes(territory) && localPopulation < 2000) {
      return { mode: "vhd", minutes: 30, summary: "VHD do 30 minut" };
    }
  } else if (item.id === "VYB_02B") {
    if (["A", "B", "C"].includes(territory) && localPopulation > 5000) {
      return { mode: "distance", distance: 800, density: standard.densityByDistance[800], summary: "skutečná pěší vzdálenost 800 m" };
    }
    if (["B", "C"].includes(territory) && localPopulation < 5000) {
      return { mode: "vhd", minutes: 30, summary: "VHD do 30 minut" };
    }
  }
  return {
    mode: "manual",
    summary: "Tabulka MMR pro tuto kombinaci typu území a počtu obyvatel nestanoví číselný limit; dostupnost ověřte podle skutečné sítě."
  };
}

function getMmrSchoolRecommendation(item) {
  const standard = MMR_SCHOOL_STANDARDS[item.id];
  if (!standard) return null;
  const derived = derivedValues();
  const localPopulation = numericOrNull(derived.O_uzemi_cil);
  const projectPopulation = numericOrNull(getPrimary("O_celk"));
  const catchmentPopulation = numericOrNull(derived.O_prah);
  const totalDemand = projectPopulation === null
    ? null : Math.ceil(projectPopulation / standard.populationPerSeat);
  const accessRule = getMmrSchoolAccessRule(item);
  const standardApplicable = localPopulation === null || accessRule.mode !== "manual";
  const capacity = getCapacityEdit(item);
  const capacityVerified = capacity.satisfiesNearbyCapacity === true;
  const freeCapacity = numericOrNull(capacity.quantity);
  const source = String(capacity.source ?? "").trim();
  const spatialEstimate = accessRule.mode === "distance"
    ? calculateSpatialEstimate(item) : null;
  const areaHa = numericOrNull(getPrimary(getSecondary("Plocha pro odhad dostupnosti") || "S_zastav"));
  const density = localPopulation !== null && areaHa > 0
    ? localPopulation / areaHa : null;
  const densityMeets = accessRule.density === undefined || density === null
    ? null : density >= accessRule.density;
  const capacityTier = localPopulation === null ? null
    : standard.capacityTiers?.filter((tier) => localPopulation >= tier.population).at(-1) ?? null;
  const siteMinimumCapacity = capacityTier?.minimumSeats ?? standard.minimumCapacity;
  const siteMaximumCapacity = capacityTier?.maximumSeats ?? standard.maximumCapacity ?? Infinity;
  const populationThresholdCapacity = !standardApplicable || localPopulation === null ? 0
    : item.id === "VYB_01"
      ? localPopulation >= standard.minimumPopulation ? standard.minimumCapacity : 0
      : item.id === "VYB_02A"
        ? localPopulation >= standard.minimumPopulation ? siteMinimumCapacity
          : localPopulation >= standard.smallSchoolPopulation ? standard.smallSchoolCapacity : 0
        : localPopulation >= standard.minimumPopulation ? siteMinimumCapacity : 0;
  const spatialCapacityFloor = standardApplicable && densityMeets === true && Number.isFinite(spatialEstimate?.count)
    ? spatialEstimate.count * siteMinimumCapacity : 0;
  const capacityDemand = !standardApplicable ? 0 : totalDemand === null ? null
    : Math.max(totalDemand, populationThresholdCapacity);
  const appliedCapacity = capacityVerified ? (totalDemand ?? 0) : 0;
  const requiredCapacity = capacityDemand === null ? null
    : !standardApplicable || capacityVerified ? 0 : Math.max(capacityDemand, spatialCapacityFloor);
  const minimumFacilityCount = !standardApplicable || capacityDemand === null || capacityVerified ? 0
    : Math.max(
      Math.ceil(capacityDemand / siteMaximumCapacity),
      densityMeets === true ? spatialEstimate?.count ?? 0 : 0
    );
  const sourcePopulation = numericOrNull(state.defaults.primary.O_celk)
    ?? numericOrNull(getPrimary("O_celk"));
  const modelReferenceDemand = sourcePopulation === null
    ? numericOrNull(item.demand)
    : Math.ceil(sourcePopulation / standard.populationPerSeat);
  const accessStatus = !standardApplicable ? "Mimo podmínky standardu MMR"
    : capacityVerified ? "Vyhovuje – kapacita okolí potvrzena" : "Neověřeno uživatelem";
  return {
    ...standard,
    projectPopulation,
    localPopulation,
    catchmentPopulation,
    totalDemand,
    capacityDemand,
    freeCapacity,
    appliedCapacity,
    requiredCapacity,
    spatialPointCount: spatialEstimate?.count ?? null,
    spatialCapacityFloor,
    capacityVerified,
    accessRule,
    standardApplicable,
    accessStatus,
    source,
    measuredAccess: null,
    density,
    densityMeets,
    areaHa,
    modelReferenceDemand,
    populationThresholdCapacity,
    capacityTier,
    siteMinimumCapacity,
    siteMaximumCapacity,
    minimumFacilityCount,
    minimumPopulationReached: localPopulation !== null
      && localPopulation >= standard.minimumPopulation
  };
}

function getMmrSchoolSummaryText(item) {
  const recommendation = getMmrSchoolRecommendation(item);
  if (!recommendation || recommendation.totalDemand === null) return "chybí populační vstup";
  if (!recommendation.standardApplicable) {
    return `${formatNumber(recommendation.totalDemand)} ${text(getCapacityResultUnit(item), "míst")} modelová potřeba · mimo práh MMR`;
  }
  const quantity = recommendation.requiredCapacity;
  return `${formatNumber(quantity)} ${text(getCapacityResultUnit(item), "míst")}`;
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
  let volumePerFloorArea = input(item.id === "VYB_17B"
    ? "MODEL_CONSTRUCTION_HEIGHT" : "MODEL_VOLUME_PER_FLOOR_M2");
  const rate = getTargetedPlanCostRate(item);
  const surcharge = input("MODEL_COST_SURCHARGE");
  if ([capacity, areaPerUnit, volumePerFloorArea, rate, surcharge]
    .some((value) => !Number.isFinite(value))) return null;

  if (item.id === "VYB_09D") {
    const vehicles = capacity;
    const vehiclesPerBase = input("MODEL_VEHICLES_PER_BASE");
    if (![vehiclesPerBase, input("MODEL_BASE_AREA"), sitePerUnit,
      input("MODEL_PAVED_PER_BASE"), input("MODEL_PAVED_PER_UNIT")].every(Number.isFinite)) return null;
    const bases = Math.ceil(vehicles / vehiclesPerBase);
    units = bases;
    usableArea = bases * input("MODEL_BASE_AREA") + vehicles * areaPerUnit;
    totalFloorArea = usableArea;
    siteArea = bases * sitePerUnit;
    pavedArea = bases * input("MODEL_PAVED_PER_BASE") + vehicles * input("MODEL_PAVED_PER_UNIT");
    unitLabel = "výjezdových základen";
    capacity = vehicles;
    detail = `${formatNumber(vehicles)} výjezdových vozů; ${formatNumber(vehiclesPerBase)} vozů/základnu`;
  } else if (["VYB_16A", VOLUNTEER_FIRE_STATION_ID].includes(item.id)) {
    const vehicles = capacity;
    const vehiclesPerBase = input("MODEL_VEHICLES_PER_BASE");
    if (![vehiclesPerBase, input("MODEL_BASE_AREA"), sitePerUnit,
      input("MODEL_PAVED_PER_BASE"), input("MODEL_PAVED_PER_UNIT")].every(Number.isFinite)) return null;
    const stations = Math.ceil(vehicles / vehiclesPerBase);
    units = stations;
    usableArea = stations * input("MODEL_BASE_AREA") + vehicles * areaPerUnit;
    totalFloorArea = usableArea;
    siteArea = stations * sitePerUnit;
    pavedArea = stations * input("MODEL_PAVED_PER_BASE") + vehicles * input("MODEL_PAVED_PER_UNIT");
    unitLabel = item.id === VOLUNTEER_FIRE_STATION_ID
      ? "hasičských zbrojnic JSDH" : "hasičských stanic";
    capacity = vehicles;
    detail = `${formatNumber(vehicles)} požárních vozidel; ${formatNumber(vehiclesPerBase)} vozidel/stanici`;
  } else if (item.id === "VYB_17A") {
    const officers = capacity;
    const officersPerStation = input("MODEL_CAPACITY_PER_FACILITY");
    if (![officersPerStation, input("MODEL_BASE_AREA"), areaPerUnit, sitePerUnit,
      input("MODEL_PAVED_PER_BASE")].every(Number.isFinite)) return null;
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
    if (![floorFactor, sitePerUnit, pavedShare, constructionHeight].every(Number.isFinite)) return null;
    usableArea = people * areaPerUnit;
    totalFloorArea = usableArea * floorFactor;
    siteArea = people * sitePerUnit;
    pavedArea = siteArea * pavedShare;
    unitLabel = "osob";
    detail = `${formatNumber(people)} osob × ${formatNumber(areaPerUnit)} m²/osobu`;
    volumePerFloorArea = constructionHeight;
  } else {
    if (![floorFactor, sitePerUnit, pavedShare].every(Number.isFinite)) return null;
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
  const count = ["VYB_09D", "VYB_16A", VOLUNTEER_FIRE_STATION_ID, "VYB_17A"].includes(item.id)
    ? units : Math.ceil(capacity / input(item.id === "VYB_01_J"
      ? "MODEL_GROUP_CAPACITY" : "MODEL_CAPACITY_PER_FACILITY"));
  const countForms = {
    VYB_01_J: ["skupina", "skupiny", "skupin"],
    VYB_02C_1: ["škola", "školy", "škol"],
    VYB_02C_2: ["škola", "školy", "škol"],
    VYB_09A_4: ["ordinace", "ordinace", "ordinací"],
    VYB_17: ["ordinace", "ordinace", "ordinací"],
    VYB_18: ["ordinace", "ordinace", "ordinací"],
    VYB_19: ["ordinace", "ordinace", "ordinací"],
    VYB_09B_1: ["ambulance", "ambulance", "ambulancí"],
    VYB_09B_2: ["ambulance", "ambulance", "ambulancí"],
    VYB_09D: ["základna", "základny", "základen"],
    VYB_16A: ["stanice", "stanice", "stanic"],
    [VOLUNTEER_FIRE_STATION_ID]: ["zbrojnice JSDH", "zbrojnice JSDH", "zbrojnic JSDH"],
    VYB_17A: ["služebna", "služebny", "služeben"],
    VYB_17B: ["úkryt", "úkryty", "úkrytů"]
  }[item.id];
  const capacityForms = {
    VYB_01_J: ["místo", "místa", "míst"],
    VYB_02C_1: ["žákovské místo", "žákovská místa", "žákovských míst"],
    VYB_02C_2: ["žákovské místo", "žákovská místa", "žákovských míst"],
    VYB_17: ["jednotka souč. zátěže", "jednotky souč. zátěže", "jednotek souč. zátěže"],
    VYB_18: ["jednotka souč. zátěže", "jednotky souč. zátěže", "jednotek souč. zátěže"],
    VYB_19: ["jednotka souč. zátěže", "jednotky souč. zátěže", "jednotek souč. zátěže"],
    VYB_09B_1: ["ambulance", "ambulance", "ambulancí"],
    VYB_09B_2: ["ambulance", "ambulance", "ambulancí"],
    VYB_09D: ["výjezdový vůz", "výjezdové vozy", "výjezdových vozů"],
    VYB_16A: ["požární vozidlo", "požární vozidla", "požárních vozidel"],
    [VOLUNTEER_FIRE_STATION_ID]: ["požární vozidlo", "požární vozidla", "požárních vozidel"],
    VYB_17A: ["policista", "policisté", "policistů"],
    VYB_17B: ["osoba", "osoby", "osob"]
  }[item.id];
  const customRate = Object.hasOwn(state.itemEstimateValues[item.key] || {}, "MODEL_COST_PER_M3");
  const capacityDisplay = item.id === "VYB_09A_4"
    ? `${formatNumber(capacity)} souč. zátěže`
    : `${formatNumber(capacity)} ${czechCountUnit(capacity, ...capacityForms)}`;
  const countDisplay = `${formatNumber(count)} ${czechCountUnit(count, ...countForms)}`;
  const countMetricLabel = {
    VYB_01_J: "Modelový počet dětských skupin",
    VYB_02C_1: "Modelový počet školních zařízení",
    VYB_02C_2: "Modelový počet ZUŠ",
    VYB_09A_4: "Modelový počet ordinačních modulů",
    VYB_17: "Modelový počet ordinací",
    VYB_18: "Modelový počet ordinací",
    VYB_19: "Modelový počet ordinací",
    VYB_09B_1: "Modelový počet odborných ambulancí",
    VYB_09B_2: "Modelový počet odborných ambulancí",
    VYB_09D: "Modelový počet výjezdových základen",
    VYB_16A: "Modelový počet hasičských stanic",
    [VOLUNTEER_FIRE_STATION_ID]: "Modelový počet hasičských zbrojnic JSDH",
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
    usableArea,
    totalFloorArea,
    volume,
    cost,
    count,
    capacity,
    siteArea,
    footprintArea,
    pavedArea,
    source: `${customRate ? `Uživatelsky upravená sazba (výchozí reference: ${assumptions.rateSource})` : assumptions.rateSource}${getRtsBuildingGroup(item)
      ? `; ${RTS_801_SOURCE} JKSO ${getRtsBuildingGroup(item)}, konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}, ${customRate ? "výchozí " : ""}sazba ${formatNumber(getRtsConstructionRate(item))} Kč/m³ OP`
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

function getTargetedPlanEstimate(item, includeSpatialFloor = false) {
  const context = getTargetedPlanContext();
  const { projectPopulation, catchmentPopulation, territory } = context;
  const estimate = (metrics, equations, sources, facilityNeed = null, modelEquations = null) => {
    const coverage = includeSpatialFloor && hasSpatialPointFloor(item)
      ? calculateSpatialEstimate(item) : null;
    const capacityPerFacility = getTargetedPlanInput(item, "MODEL_CAPACITY_PER_FACILITY");
    const spatialCapacityFloor = Number.isFinite(coverage?.count)
      && capacityPerFacility > 0 && facilityNeed !== null
      && facilityNeed !== undefined && facilityNeed > 0
      ? coverage.count * capacityPerFacility : null;
    const plannedFacilityNeed = facilityNeed === null || facilityNeed === undefined
      ? facilityNeed : Math.max(facilityNeed, spatialCapacityFloor ?? 0);
    if (spatialCapacityFloor !== null) {
      const needMetricLabels = {
        VYB_09A_4: "Poptávka po gynekologii po odečtu ověřené kapacity",
        VYB_17: "Poptávka nepokrytá ověřenou kapacitou",
        VYB_18: "Poptávka nepokrytá ověřenou kapacitou",
        VYB_19: "Poptávka nepokrytá ověřenou kapacitou",
        VYB_17B: "Poptávka nepokrytá ověřenou kapacitou úkrytů"
      };
      const needMetric = metrics.find((metric) => metric.label === needMetricLabels[item.id]);
      if (needMetric && plannedFacilityNeed !== null) needMetric.value = plannedFacilityNeed;
      const countMetric = metrics.find((metric) => metric.label.startsWith("Modelový počet "));
      if (countMetric && plannedFacilityNeed !== null) {
        countMetric.value = Math.ceil(plannedFacilityNeed / capacityPerFacility);
      }
      metrics.push(
        { label: "Orientační počet bodů pro plošné pokrytí", value: coverage.count, unit: "bodů" },
        { label: "Kapacitní minimum pro plošné pokrytí", value: spatialCapacityFloor, unit: "kapacitních jednotek" }
      );
      equations.push(
        `Orientační návrh po plošném pokrytí = max(${formatNumber(facilityNeed)} kapacitních jednotek podle poptávky, ${formatNumber(coverage.count)} bodů × ${formatNumber(capacityPerFacility)} kapacitních jednotek/bod) = ${formatNumber(plannedFacilityNeed)} kapacitních jednotek.`
      );
      sources.push("Plošný počet bodů je konzervativní geometrický odhad z MMR docházkového dosahu. Předpokládá, že stávající zařízení zatím nepokrývají celé řešené území; jde o návrhové minimum k ověření v GIS, nikoli o právní kvótu.");
    }
    const facilityPlan = getTargetedFacilityPlan(item, plannedFacilityNeed);
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
      verificationEquations, sources, facilityPlan,
      spatialPointCount: spatialCapacityFloor !== null ? coverage.count : null
    };
  };

  if (isHealthcareDemandItem(item) && item.id !== "VYB_09A_4") {
    const demographicCoefficient = getTargetedPlanInput(item, "MODEL_DEMOGRAPHIC_COEFFICIENT");
    const populationPerFte = getDoctorPopulationPerFte(item, includeSpatialFloor);
    const capacityPerFacility = getTargetedPlanInput(item, "MODEL_CAPACITY_PER_FACILITY");
    const population = numericOrNull(getPrimary("O_celk"));
    const eligiblePopulation = population === null || demographicCoefficient === null
      ? null : population * demographicCoefficient;
    const specialties = HEALTH_SPECIALTIES[item.id] || [];
    const demandBySpecialty = specialties.map((specialty) => ({
      profession: specialty.profession,
      demand: eligiblePopulation === null || !populationPerFte
        ? null : eligiblePopulation / populationPerFte
    }));
    const requiredFte = demandBySpecialty.some((entry) => entry.demand === null)
      ? null : demandBySpecialty.reduce((sum, entry) => sum + entry.demand, 0);
    const capacityEdit = getCapacityEdit(item);
    const availableFte = numericOrNull(capacityEdit.quantity);
    const nearbyCapacityConfirmed = capacityEdit.satisfiesNearbyCapacity === true;
    const checksIndividualAccess = ["VYB_17", "VYB_18", "VYB_19"].includes(item.id);
    const physicalAccess = ["A", "B", "C"].includes(territory);
    const municipalityThresholdMet = context.projectPopulation !== null
      && context.projectPopulation > 2000;
    const walkDistance = checksIndividualAccess
      ? getTargetedPlanInput(item, "HEALTHCARE_WALK_DISTANCE") : null;
    const driveMinutes = checksIndividualAccess
      ? getTargetedPlanInput(item, "HEALTHCARE_DRIVE_MINUTES") : null;
    const accessRequired = includeSpatialFloor && checksIndividualAccess
      && (physicalAccess ? municipalityThresholdMet : territory === "D");
    const accessVerified = accessRequired
      && (physicalAccess ? walkDistance !== null : driveMinutes !== null);
    const accessPasses = !accessRequired || accessVerified
      && (physicalAccess ? walkDistance <= 600 : driveMinutes <= 35);
    const accessStatus = !accessRequired ? "Standard se na tuto kombinaci nevztahuje"
      : !accessVerified ? "Neověřeno" : accessPasses ? "Vyhovuje" : "Nevyhovuje";
    const capacityAccessible = !includeSpatialFloor || !accessRequired
      || accessVerified && accessPasses;
    const nearbyCapacityApplied = nearbyCapacityConfirmed && capacityAccessible;
    const minimumPopulation = includeSpatialFloor
      ? getTargetedPlanInput(item, "MODEL_MINIMUM_POPULATION_PER_FTE") : null;
    const minimumDensity = includeSpatialFloor
      ? getTargetedPlanInput(item, "MODEL_MINIMUM_DENSITY_PER_HA") : null;
    const twoDayPopulationMet = minimumPopulation !== null && context.projectPopulation !== null
      && context.projectPopulation >= minimumPopulation;
    const twoDayDensityMet = minimumDensity !== null && context.populationDensity !== null
      && context.populationDensity >= minimumDensity;
    const remainingFte = requiredFte === null ? null
      : !capacityAccessible ? requiredFte
        : nearbyCapacityApplied ? 0 : Math.max(0, requiredFte - (availableFte ?? 0));
    const capacityVerified = capacityAccessible && (availableFte !== null || nearbyCapacityConfirmed);
    const clinics = remainingFte === null || remainingFte === 0 || !capacityPerFacility
      ? remainingFte === 0 ? 0 : null
      : Math.ceil(remainingFte / capacityPerFacility);
    const isSpecialistGroup = item.id === "VYB_09B_1" || item.id === "VYB_09B_2";
    const hasMmrPopulationRatio = (TARGETED_PLAN_INPUTS[item.id] || [])
      .some((field) => field.code === "MODEL_POPULATION_PER_FTE");
    const populationRatioUnit = includeSpatialFloor && hasMmrPopulationRatio
      ? isSpecialistGroup ? "obyv./ambulanci" : "obyvatel/lékaře"
      : "pacientů/lékaře";
    const unitLabel = isSpecialistGroup ? "ambulancí" : "ordinací";
    const populationLabel = "Demograficky započtená populace z přírůstku záměru";
    const metrics = [
      { label: "Přírůstek obyvatel řešeného území (záměr)", value: population, unit: "obyvatel" },
      { label: "Demografický koeficient", value: demographicCoefficient, unit: "podíl populace" },
      { label: populationLabel, value: eligiblePopulation, unit: "obyvatel" },
      { label: includeSpatialFloor && hasMmrPopulationRatio
        ? "Počet obyvatel na lékaře dle MMR"
        : includeSpatialFloor
          ? "Poměr pro modelový odhad kapacity (MMR jej nestanoví)"
        : "Registrovaní pacienti na lékaře při plném úvazku (státní pracovní základ)",
        value: populationPerFte, unit: populationRatioUnit },
      ...(minimumPopulation === null ? [] : [
        { label: "Minimální populace pro ordinaci 2 dny týdně (MMR)", value: minimumPopulation, unit: "obyvatel" },
        { label: "Minimální hustota pro ordinaci 2 dny týdně (MMR)", value: minimumDensity, unit: "obyvatel/ha" },
        { label: "Splněn populační a hustotní práh pro ordinaci 2 dny týdně",
          value: twoDayPopulationMet && twoDayDensityMet ? "Ano" : "Ne", unit: "" }
      ]),
      ...demandBySpecialty.map((entry) => ({
        label: `Modelová potřeba – ${entry.profession}`,
        value: entry.demand,
        unit: isSpecialistGroup ? "ambulance" : "lékařského úvazku"
      })),
      { label: "Požadovaná modelová potřeba", value: requiredFte, unit: isSpecialistGroup ? "ambulancí" : "lékařských úvazků" },
      { label: "Ověřená volná kapacita v okolí", value: availableFte, unit: isSpecialistGroup ? "ambulancí" : "lékařských úvazků" },
      { label: "Kapacita okolí potvrzena uživatelem", value: nearbyCapacityConfirmed ? "Ano" : "Ne", unit: "" },
      ...(includeSpatialFloor && checksIndividualAccess ? [
        { label: "Kontrola dostupnosti MMR", value: accessStatus, unit: "" },
        { label: "Ověřená pěší vzdálenost", value: physicalAccess ? walkDistance : null, unit: "m" },
        { label: "Ověřený dojezd (typ D)", value: physicalAccess ? null : driveMinutes, unit: "min" }
      ] : []),
      { label: "Poptávka nepokrytá ověřenou kapacitou", value: remainingFte, unit: isSpecialistGroup ? "ambulancí" : "lékařských úvazků" },
      { label: `Modelový počet ${unitLabel}`, value: clinics, unit: unitLabel }
    ];
    const equations = [
      population === null || eligiblePopulation === null
        ? "Chybí počet obyvatel nebo demografický koeficient; nelze vypočítat započtenou populaci."
        : `Započtená populace záměru = ${formatNumber(population)} nových obyvatel × ${formatNumber(demographicCoefficient * 100)} % = ${formatNumber(eligiblePopulation)} obyvatel.`,
      requiredFte === null || !populationPerFte
        ? "Chybí populační poměr na lékařskou souč. zátěž; poptávku nelze vypočítat."
        : isSpecialistGroup
          ? `Poptávka odborných ambulancí = ${formatNumber(eligiblePopulation)} obyvatel přírůstku ÷ ${formatNumber(populationPerFte)} ${populationRatioUnit} × ${specialties.length} odborností = ${formatNumber(requiredFte)} ambulancí.`
          : `Poptávka ordinací = ${formatNumber(eligiblePopulation)} obyvatel přírůstku ÷ ${formatNumber(populationPerFte)} ${populationRatioUnit} = ${formatNumber(requiredFte)} lékařských úvazků.`,
      ...(minimumPopulation === null ? [] : [
        `Kontrola MMR pro ordinaci dva dny týdně: ${formatNumber(context.projectPopulation)} obyvatel ${twoDayPopulationMet ? "splňuje" : "nesplňuje"} minimum ${formatNumber(minimumPopulation)}; hustota ${formatNumber(context.populationDensity)} obyv./ha ${twoDayDensityMet ? "splňuje" : "nesplňuje"} minimum ${formatNumber(minimumDensity)} obyv./ha. Jde o samostatný práh, nikoli náhradu výpočtu potřeby z přírůstku.`
      ]),
      nearbyCapacityApplied
        ? "Uživatel potvrdil, že okolní kapacita splňuje potřebu; nová ordinace se proto nenavrhuje."
        : !capacityAccessible
          ? `Kapacita se neodečítá, protože dostupnost MMR je ${accessStatus.toLowerCase()}; ověřte konkrétní pěší trasu do 600 m nebo dojezd typu D do 35 min.`
          : `Po odečtu ověřené volné kapacity = max(0, ${formatNumber(requiredFte)} − ${capacityVerified ? formatNumber(availableFte) : "0 (neověřeno)"}) = ${formatNumber(remainingFte)} ${isSpecialistGroup ? "ambulancí" : "lékařských úvazků"}; modelový počet zařízení = ceil(${formatNumber(remainingFte)} ÷ ${formatNumber(capacityPerFacility)}) = ${formatNumber(clinics)} ${unitLabel}.`
    ];
    const source = item.id === "VYB_09B_1" || item.id === "VYB_09B_2"
      ? includeSpatialFloor
        ? "MMR určuje sídelní přítomnost a dostupnost těchto skupin, nikoli populační kvótu ambulance. Počet ambulancí je samostatný plánovací odhad založený na editovatelném vstupu."
        : "Zákon neurčuje jednotnou kapacitu ambulance. Výpočet používá editovatelný obecný pracovní základ 1 700 registrovaných pacientů na lékaře při plném úvazku (rozmezí 1 600–1 800); místní oborová skladba a registrace vyžadují ověření."
      : item.id === "VYB_17"
        ? includeSpatialFloor
          ? "Doporučený poměr vychází z MMR, s. 27; zákonný výpočet používá samostatný, editovatelný počet registrovaných pacientů na plný lékařský úvazek."
          : "Zákon neurčuje jednotnou kapacitu. Používá se editovatelný obecný pracovní základ 1 700 registrovaných pacientů na lékaře při plném úvazku (rozmezí 1 600–1 800)."
        : item.id === "VYB_18"
          ? includeSpatialFloor
            ? "Doporučený poměr vychází z MMR, s. 27; zákonný výpočet používá samostatný, editovatelný počet registrovaných pacientů na plný lékařský úvazek."
            : "Zákon neurčuje jednotnou kapacitu. Používá se editovatelný obecný pracovní základ 1 700 registrovaných pacientů na lékaře při plném úvazku (rozmezí 1 600–1 800)."
          : item.id === "VYB_19"
            ? includeSpatialFloor
              ? "Doporučený poměr vychází z MMR, s. 27; zákonný výpočet používá samostatný, editovatelný počet registrovaných pacientů na plný lékařský úvazek."
              : "Zákon neurčuje jednotnou kapacitu. Používá se editovatelný obecný pracovní základ 1 700 registrovaných pacientů na lékaře při plném úvazku (rozmezí 1 600–1 800)."
            : "Kapacitní poměr je pouze scénářový předpoklad; ověřte jej podle lokálních registrací poskytovatelů.";
    return estimate(metrics, equations, [
      `${source} Počet oborů ve skupině vychází z evidovaných odborností v kartě. Kapacita ordinace je editovatelný předpoklad; volná kapacita v okolí se odečte jen po zadání její hodnoty.`
    ], remainingFte);
  }

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
      && existingDistance !== null && existingDistance <= coverageDistance;
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
    const studentShare = getTargetedPlanInput(item, "MODEL_DEMOGRAPHIC_SHARE");
    const seatsPerSchool = getTargetedPlanInput(item, "MODEL_CAPACITY_PER_FACILITY");
    const localPopulation = numericOrNull(derivedValues().O_uzemi_cil);
    const projectIncrement = numericOrNull(getPrimary("O_celk"));
    const isApplicable = secondarySchool
      ? ["A", "B"].includes(territory)
      : territory === "A"
        || ["B", "C"].includes(territory) && localPopulation > 5000;
    const populationThresholdReached = secondarySchool
      ? ["A", "B"].includes(territory) && localPopulation !== null && localPopulation > 20000
      : localPopulation !== null && (territory === "A"
        || ["B", "C"].includes(territory) && localPopulation > 5000);
    const localPresenceRecommended = isMmrSchoolPresenceRecommended(item, localPopulation) ?? false;
    const capacity = getCapacityEdit(item);
    const capacityVerified = capacity.satisfiesNearbyCapacity === true;
    const studentDemand = projectIncrement !== null && studentShare !== null
      ? Math.ceil(projectIncrement * studentShare) : null;
    const unmetSeats = studentDemand === null ? null
      : capacityVerified ? 0 : studentDemand;
    const newSchools = unmetSeats === null || !Number.isFinite(seatsPerSchool) || seatsPerSchool <= 0
      ? null : Math.ceil(unmetSeats / seatsPerSchool);
    const metrics = [
      { label: "Současní obyvatelé obce + přírůstek záměru (základ prahu MMR)",
        value: localPopulation, unit: "obyvatel" },
      { label: secondarySchool
        ? "Prahová hodnota doporučené přítomnosti SŠ (> 20 000 obyvatel v A/B)"
        : "Typová/populační podmínka přítomnosti ZUŠ (A; B/C > 5 000 obyvatel)",
        value: populationThresholdReached ? "Splněna" : "Nesplněna", unit: "" },
      { label: "Doporučená sídelní přítomnost MMR", value: localPresenceRecommended ? "Ano" : "Ne", unit: "" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Požadavek MMR použitelný pro typ území", value: isApplicable ? "Ano" : "Ne – položka není předmětem standardu", unit: "" }
    ];
    metrics.push(
      { label: secondarySchool
        ? `Scénářový podíl nové populace v cílové středoškolské kohortě (${formatNumber((studentShare ?? 0) * 100)} %)`
        : `Scénářový podíl nové populace využívající ZUŠ (${formatNumber((studentShare ?? 0) * 100)} %)`,
        value: studentShare === null ? null : studentShare * 100, unit: "%" },
      { label: "Modelová poptávka z přírůstku obyvatel", value: studentDemand, unit: "žáků" },
      { label: "Kapacita okolí potvrzena uživatelem", value: capacityVerified ? "Ano" : "Neověřeno", unit: "" },
      { label: "Nová kapacita k zajištění", value: unmetSeats, unit: "míst" },
      { label: "Počet modelových zařízení", value: newSchools, unit: "škol" }
    );
    return estimate(metrics, [
      studentDemand === null
        ? "Chybí počet obyvatel nebo demografický podíl; poptávku po místech nelze vypočítat."
        : `Poptávka po nových místech = ceil(${formatNumber(projectIncrement)} nových obyvatel × ${formatNumber(studentShare * 100)} %) = ${formatNumber(studentDemand)} žáků; současní obyvatelé se započítávají pouze do prahů sídelní přítomnosti.`,
      secondarySchool
        ? `MMR: v typech A/B platí dostupnost VHD do 45 minut. Doporučená sídelní přítomnost SŠ = ${formatNumber(numericOrNull(getPrimary("O_sidlo")))} současných obyvatel + ${formatNumber(projectIncrement)} obyvatel přírůstku = ${formatNumber(localPopulation)}; práh je více než 20 000. Typy C/D nejsou tímto standardem pokryty.`
        : `MMR: sídelní přítomnost ZUŠ je stanovena pro typ A a doporučena pro B/C při více než 5 000 obyvatelích. Prah = ${formatNumber(numericOrNull(getPrimary("O_sidlo")))} současných obyvatel + ${formatNumber(projectIncrement)} obyvatel přírůstku = ${formatNumber(localPopulation)}; typ D a B/C do 5 000 obyvatel nemají standard. Pěší limity se nepoužívají.`,
      capacityVerified
        ? "Uživatel potvrdil, že stávající okolní kapacita pokrývá návrhovou potřebu; nová kapacita se proto navrhuje jako 0."
        : `Bez potvrzení kapacity okolí zůstává k zajištění celá nová poptávka ${formatNumber(studentDemand)} míst, tedy ceil(${formatNumber(studentDemand)} ÷ ${formatNumber(seatsPerSchool)}) = ${newSchools} modelových škol.`,
      `${studentShare === null ? "Poptávkový podíl není vyplněn." : `Poptávkový podíl ${formatNumber(studentShare * 100)} %`}${seatsPerSchool === null ? " Kapacita zařízení není vyplněna." : ` a kapacita ${formatNumber(seatsPerSchool)} míst na zařízení`} jsou editovatelné projektové předpoklady.`
    ], [
      `TB050MMR01, tabulka 1, tištěná s. 8; prahy SŠ nad 20 000 a ZUŠ nad 5 000 vyhodnocují současné obyvatele obce spolu s přírůstkem záměru. Poptávkový podíl a velikost typového zařízení jsou editovatelné uživatelské předpoklady.`
    ], unmetSeats === null ? null : unmetSeats, [studentDemand === null
      ? "Chybí počet obyvatel nebo demografický podíl; poptávku nelze vypočítat."
      : `Poptávka po místech = ceil(${formatNumber(projectIncrement)} nových obyvatel × ${formatNumber(studentShare * 100)} %) = ${formatNumber(studentDemand)} žáků.`]);
  }

  if (item.id === "VYB_09A_4") {
    const residentsPerDoctor = getDoctorPopulationPerFte(item, includeSpatialFloor);
    const residentsPerDoctorUnit = includeSpatialFloor
      ? "obyvatel/lékaře" : "registrovaných pacientů/lékaře na plný úvazek";
    const demographicCoefficient = getTargetedPlanInput(item, "MODEL_DEMOGRAPHIC_COEFFICIENT");
    const population = numericOrNull(getPrimary("O_celk"));
    const eligiblePopulation = population !== null && demographicCoefficient !== null
      ? population * demographicCoefficient : null;
    const fte = eligiblePopulation !== null && residentsPerDoctor
      ? eligiblePopulation / residentsPerDoctor : null;
    const minimumPopulation = includeSpatialFloor
      ? getTargetedPlanInput(item, "MODEL_MINIMUM_POPULATION_PER_FTE") : null;
    const minimumDensity = includeSpatialFloor
      ? getTargetedPlanInput(item, "MODEL_MINIMUM_DENSITY_PER_HA") : null;
    const twoDayPopulationMet = context.projectPopulation !== null
      && minimumPopulation !== null && context.projectPopulation >= minimumPopulation;
    const twoDayDensityMet = context.populationDensity !== null
      && minimumDensity !== null && context.populationDensity >= minimumDensity;
    const capacityEdit = getCapacityEdit(item);
    const availableFte = numericOrNull(capacityEdit.quantity)
      ?? getTargetedPlanInput(item, "GYNECOLOGY_AVAILABLE_FTE");
    const walkDistance = getTargetedPlanInput(item, "GYNECOLOGY_WALK_DISTANCE");
    const driveMinutes = getTargetedPlanInput(item, "GYNECOLOGY_DRIVE_MINUTES");
    const accessSource = getTargetedPlanSource(item, "GYNECOLOGY_ACCESS_SOURCE");
    const source = accessSource || String(capacityEdit.source ?? "").trim();
    const physicalAccess = ["A", "B", "C"].includes(territory);
    const municipalityThresholdMet = physicalAccess
      && projectPopulation !== null && projectPopulation > 2000;
    const standardApplicable = !includeSpatialFloor || !physicalAccess || municipalityThresholdMet;
    const coverage = includeSpatialFloor && municipalityThresholdMet
      ? getTargetedCoverageEstimate(600) : null;
    const accessVerified = includeSpatialFloor && (
      physicalAccess ? walkDistance !== null : driveMinutes !== null);
    const accessPasses = !includeSpatialFloor || accessVerified && standardApplicable && (
      physicalAccess ? walkDistance <= 600 : driveMinutes <= 35);
    const nearbyCapacityConfirmed = capacityEdit.satisfiesNearbyCapacity === true;
    const nearbyCapacityApplied = nearbyCapacityConfirmed
      && (!includeSpatialFloor || !standardApplicable || accessVerified && accessPasses);
    const capacityVerified = nearbyCapacityApplied || availableFte !== null
      && (includeSpatialFloor ? accessVerified && accessPasses : true);
    const remainingFte = fte === null ? null : !standardApplicable || nearbyCapacityApplied
      ? 0 : capacityVerified ? Math.max(0, fte - availableFte) : fte;
    return estimate([
      { label: "Přírůstek obyvatel řešeného území (záměr)", value: population, unit: "obyvatel" },
      { label: "Demografický koeficient populace pacientek", value: demographicCoefficient, unit: "podíl populace" },
      { label: "Demograficky započtená populace z přírůstku záměru", value: eligiblePopulation, unit: "obyvatel" },
      { label: includeSpatialFloor
        ? "Optimální počet obyvatel na gynekologa (MMR 2020)"
        : "Registrovaní pacienti na gynekologa při plném úvazku (státní pracovní základ)",
        value: residentsPerDoctor, unit: residentsPerDoctorUnit },
      ...(includeSpatialFloor ? [
        { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
        { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" }
      ] : []),
      { label: "Typ území", value: territory, unit: "" },
      ...(includeSpatialFloor ? [
      { label: "Prahová přítomnost v obci (> 2 000; A/B/C)", value: physicalAccess
        ? municipalityThresholdMet ? "Dosažen" : "Nedosažen – doporučení se neuplatní"
        : "Nepoužije se pro typ D", unit: "" },
      { label: "Minimální populace pro ordinaci 2 dny týdně (MMR)", value: minimumPopulation, unit: "obyvatel" },
      { label: "Minimální hustota pro ordinaci 2 dny týdně (MMR)", value: minimumDensity, unit: "obyvatel/ha" },
      { label: "Splněn populační práh ordinace 2 dny týdně", value: twoDayPopulationMet ? "Ano" : "Ne", unit: "" },
      { label: "Splněna hustota ordinace 2 dny týdně", value: twoDayDensityMet ? "Ano" : "Ne", unit: "" },
      { label: "Standard MMR použitelný pro typ/populaci", value: standardApplicable ? "Ano" : "Ne", unit: "" },
      ] : [
        { label: "Kapacitní poměr je zákonný normativ", value: "Ne – editovatelný obecný pracovní údaj", unit: "" }
      ]),
      { label: "Indikativní počet lékařských úvazků", value: fte, unit: "lékařských úvazků" },
      { label: "Ověřená volná kapacita", value: availableFte, unit: "lékařských úvazků" },
      { label: "Kapacita okolí potvrzena uživatelem", value: nearbyCapacityConfirmed ? "Ano" : "Ne", unit: "" },
      ...(includeSpatialFloor ? [{
        label: "Potvrzení kapacity započteno po ověření dostupnosti",
        value: nearbyCapacityApplied ? "Ano" : "Ne", unit: ""
      }] : []),
      { label: "Poptávka po gynekologii po odečtu ověřené kapacity", value: remainingFte, unit: "lékařských úvazků" },
      { label: "Modelový počet ordinací", value: remainingFte === null
        ? null : Math.ceil(remainingFte / getTargetedPlanInput(item, "MODEL_CAPACITY_PER_FACILITY")), unit: "ordinací" },
      ...(includeSpatialFloor ? [{ label: "Kontrola dostupnosti", value: accessVerified
        ? accessPasses ? "Vyhovuje" : "Nevyhovuje"
        : "Neověřeno", unit: "" }] : []),
      ...(includeSpatialFloor && physicalAccess
        ? [
          { label: "Ověřená pěší vzdálenost", value: walkDistance, unit: "m" },
          { label: "Orientační body sítě při pěším dosahu 600 m", value: coverage?.count ?? null, unit: "bodů" }
        ]
        : includeSpatialFloor ? [
          { label: "Ověřený dojezd (typ D)", value: driveMinutes, unit: "min" },
          { label: "Standard dostupnosti pro typ D", value: "35 min dojezdu", unit: "" }
        ] : [])
    ], [
      population === null || demographicCoefficient === null || !residentsPerDoctor
        ? "Chybí počet obyvatel, demografický koeficient nebo populační poměr; personální indikátor nelze vypočítat."
        : `Započtená populace záměru = ${formatNumber(population)} nových obyvatel × ${formatNumber(demographicCoefficient * 100)} % = ${formatNumber(eligiblePopulation)} obyvatel; indikativní potřeba = ${formatNumber(eligiblePopulation)} ÷ ${formatNumber(residentsPerDoctor)} ${residentsPerDoctorUnit} = ${formatNumber(fte)} lékařských úvazků.`,
      ...(includeSpatialFloor ? [
      `Samostatná kontrola minimálního populačního prahu pro ordinaci dva dny týdně: ${formatNumber(context.projectPopulation)} obyvatel ${twoDayPopulationMet ? "splňuje" : "nesplňuje"} práh ${formatNumber(minimumPopulation)}; hustota ${formatNumber(context.populationDensity)} obyv./ha ${twoDayDensityMet ? "splňuje" : "nesplňuje"} práh ${formatNumber(minimumDensity)} obyv./ha. Tyto prahy nepřepočítávají poptávku z nových obyvatel.`,
      physicalAccess
        ? `Populační práh MMR pro přítomnost v obci (> 2 000 obyvatel) ${municipalityThresholdMet ? "je dosažen podle zadané populace řešeného území" : "není dosažen podle zadané populace řešeného území"}; ověřte, že tento vstup odpovídá celé obci.`
        : "Pro typ D se místo strukturálního populačního prahu uplatňuje dojezdový standard.",
      !standardApplicable
        ? "Pro typ A/B/C bez překročení 2 000 obyvatel se doporučení na sídelní přítomnost neuplatní; modelová nová potřeba ordinace je proto 0."
        : nearbyCapacityApplied
          ? "Uživatel potvrdil, že kapacita v okolí splňuje potřebu; nová ordinace se proto nenavrhuje."
          : nearbyCapacityConfirmed && includeSpatialFloor && standardApplicable
            ? "Potvrzení kapacity se nezapočítá, dokud není ověřena vyhovující pěší vzdálenost do 600 m nebo dojezd do 35 minut."
          : capacityVerified && remainingFte !== null
        ? `Poptávka po odečtu ověřené volné kapacity dostupné ordinace = max(0, ${formatNumber(fte)} − ${formatNumber(availableFte)}) = ${formatNumber(remainingFte)} lékařských úvazků.`
        : `Poptávka k vyhledání v okolí = ${formatNumber(fte)} lékařských úvazků; bez ověřené kapacity dostupné ordinace se nic neodečítá.`,
      physicalAccess
        ? coverage
          ? `Plošný odhad pokrytí: ${formatNumber(context.areaHa)} ha, typ ${territory}, 600 m pěšky; mřížka ${coverage.columns} × ${coverage.rows} = ${coverage.count} bodů.`
          : "Plošný odhad 600 m nelze vypočítat bez plochy, tvaru a koeficientu pěší sítě."
        : "Pro typ D stanoví metodika časovou dostupnost dojezdem 35 minut; nelze ji převést z geometrie plochy bez síťové analýzy.",
      accessVerified
        ? physicalAccess
          ? `Skutečná pěší vzdálenost ${formatNumber(walkDistance)} m ${accessPasses ? "splňuje" : "překračuje"} limit 600 m.`
          : `Ověřený dojezd ${formatNumber(driveMinutes)} min ${accessPasses ? "splňuje" : "překračuje"} limit 35 min.`
        : "Dostupnost poskytovatele musí být ověřena podle typu území a doložena zdrojem."
      ] : [
      "Právní rámec neurčuje populační poměr pacientů na lékaře; počet ordinací je modelován z přírůstku záměru a editovatelné standardní kapacity 1 700 registrovaných pacientů na plný úvazek."
      ]),
      `Počet lékařských úvazků se převádí na výměru ordinací podle editovatelného scénáře ${formatNumber(getTargetedPlanInput(item, "MODEL_AREA_PER_UNIT"))} m²/úvazek; výměra není právním normativem.`
    ], includeSpatialFloor ? [
      "Metodika MMR/TB050MMR01, s. 9 a 27: pro gynekologii se používá pěší dosah 600 m v typech A/B/C (sídelní přítomnost při více než 2 000 obyvatelích) a dojezd do 35 minut v typu D; optimální poměr je 14 000 obyvatel na lékaře. Na s. 27 jsou samostatně uvedeny prahy pro ordinaci dva dny týdně: 5 600 obyvatel a 84 obyvatel/ha. Všechny vstupy jsou editovatelné."
      ,
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ] : [
      "Zákon neurčuje jednotnou kapacitu gynekologické ordinace ani počet registrovaných pacientů na lékaře. Výpočet je plánovací indikátor: přírůstek záměru / editovatelný obecný pracovní základ 1 700 registrovaných pacientů na lékaře při plném úvazku (rozmezí 1 600–1 800)."
    ], remainingFte, [population === null || !residentsPerDoctor
      ? "Chybí počet obyvatel, demografický koeficient nebo počet obyvatel na lékaře; potřebu nelze vypočítat."
      : `Potřeba lékařských úvazků = (${formatNumber(population)} nových obyvatel × ${formatNumber(demographicCoefficient * 100)} %) ÷ ${formatNumber(residentsPerDoctor)} ${residentsPerDoctorUnit} = ${formatNumber(fte)} úvazků.`]);
  }

  if (item.id === "VYB_09D") {
    const covered = getTargetedPlanInput(item, "ZZS_COVERED_POPULATION");
    const responseMinutes = getTargetedPlanInput(item, "ZZS_MAX_RESPONSE_MINUTES");
    const source = getTargetedPlanSource(item, "ZZS_COVERAGE_SOURCE");
    const populationPerVehicle = getTargetedPlanInput(item, "MODEL_POPULATION_PER_VEHICLE");
    const vehiclesPerBase = getTargetedPlanInput(item, "MODEL_VEHICLES_PER_BASE");
    const validResponse = responseMinutes !== null && catchmentPopulation !== null;
    const coveredIsValid = covered !== null && catchmentPopulation !== null && covered <= catchmentPopulation;
    const uncovered = coveredIsValid ? Math.max(0, catchmentPopulation - covered) : catchmentPopulation;
    const needsBaseReview = !validResponse || responseMinutes > 20;
    const responseVehicles = catchmentPopulation === null ? null : !needsBaseReview
      ? 0 : Math.max(vehiclesPerBase, Math.ceil(Math.max(0, uncovered) / populationPerVehicle));
    const responseBases = responseVehicles === null ? null : Math.ceil(responseVehicles / vehiclesPerBase);
    return estimate([
      { label: "Obyvatelé řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Obyvatelé včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Plocha pro kontrolu dostupnosti", value: context.areaHa, unit: "ha" },
      { label: "Hustota obyvatel po záměru", value: context.populationDensity, unit: "obyvatel/ha" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Ověřená dojezdová doba ZZS (zákonná mez 20 min)", value: validResponse ? responseMinutes : null, unit: "min" },
      { label: "Ověřeně pokrytí obyvatelé", value: coveredIsValid ? covered : null, unit: "obyvatel" },
      { label: "Obyvatelé bez doloženého pokrytí", value: uncovered, unit: "obyvatel" },
      { label: `Modelová poptávka po výjezdových vozech (1 / ${formatNumber(populationPerVehicle)} obyvatel)`,
        value: responseVehicles === null ? null : `${formatNumber(responseVehicles)} ${czechCountUnit(responseVehicles, "vůz", "vozy", "vozů")}`, unit: "" },
      { label: "Modelová potřeba nových základen po kontrole dojezdu", value: responseBases, unit: "základen" },
      { label: "Výsledek podmíněného posouzení", value: needsBaseReview ? "Potřeba základny zůstává započtena" : "Nová základna se nezapočítá", unit: "" }
    ], [
      `Zákonná mez podle § 5 zákona č. 374/2011 Sb. = 20 min; ${validResponse ? `ověřeno ${formatNumber(responseMinutes)} min` : "skutečný dojezd není doložen"}.`,
      needsBaseReview
        ? `Dokud není doloženo vyhovující pokrytí do 20 minut, používá se scénářová kapacita 1 výjezdový vůz na ${formatNumber(populationPerVehicle)} obyvatel; nejméně ${formatNumber(vehiclesPerBase)} vozidla tvoří jednu základnu = ${responseVehicles} vozů.`
        : "Doložený dojezd do celého řešeného území je nejvýše 20 minut; nová základna se proto v bilanci nezapočítá, pokud kraj potvrdí pokrytí."
      ,
      coveredIsValid
        ? `Doplňková bilance obyvatel ve spádu: ${formatNumber(catchmentPopulation)} − ${formatNumber(covered)} = ${formatNumber(uncovered)} obyvatel bez doloženého pokrytí.`
        : `Bez doloženého počtu pokrytých obyvatel se za nepokrytou poptávku považuje celý spád ${formatNumber(uncovered)} obyvatel.`,
      "Plocha a typ území jsou uvedeny pro krajské/GIS posouzení; z plochy samotné nelze spolehlivě převést uliční síť na 20minutový dojezd."
    ], [
      "Zákon č. 374/2011 Sb., § 5: nejpozději do 20 minut od převzetí pokynu k výjezdu, s výjimkami stanovenými zákonem. Poměr obyvatel na vůz a počet vozů na základnu jsou editovatelné provozní scénáře, nikoli zákonná kvóta; potvrzuje je krajský plán/poskytovatel ZZS.",
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], responseVehicles, [
      catchmentPopulation === null
        ? "Chybí počet obyvatel ve spádu; počet vozů nelze vypočítat."
        : `Modelový počet vozů = max(${formatNumber(vehiclesPerBase)} vozů/základnu, ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vůz)) = ${responseVehicles} vozů při nevyhovujícím nebo neověřeném dojezdu.`
    ]);
  }

  if (item.id === VOLUNTEER_FIRE_STATION_ID) {
    const covered = getTargetedPlanInput(item, "JSDH_COVERED_POPULATION");
    const walkDistance = getTargetedPlanInput(item, "JSDH_WALK_DISTANCE");
    const availableVehicles = getTargetedPlanInput(item, "JSDH_AVAILABLE_VEHICLES");
    const populationPerVehicle = getTargetedPlanInput(item, "MODEL_POPULATION_PER_VEHICLE");
    const source = getTargetedPlanSource(item, "JSDH_COVERAGE_SOURCE");
    const accessVerified = walkDistance !== null;
    const accessPasses = accessVerified && walkDistance <= 1000;
    const coveredIsValid = accessPasses && covered !== null
      && catchmentPopulation !== null && covered <= catchmentPopulation;
    const uncovered = coveredIsValid
      ? Math.max(0, catchmentPopulation - covered) : catchmentPopulation;
    const requiredVehicles = uncovered === null
      ? null : Math.ceil(uncovered / populationPerVehicle);
    const vehicleCapacityVerified = accessPasses
      && availableVehicles !== null && Number.isInteger(availableVehicles) && availableVehicles >= 0;
    const additionalVehicles = requiredVehicles === null ? null : vehicleCapacityVerified
      ? Math.max(0, requiredVehicles - availableVehicles) : requiredVehicles;
    const needsReview = !accessPasses || !vehicleCapacityVerified || additionalVehicles > 0;
    const vehiclesForNewStations = requiredVehicles === null ? null : needsReview
      ? Math.max(1, additionalVehicles ?? 0) : 0;
    const vehiclesPerStation = getTargetedPlanInput(item, "MODEL_VEHICLES_PER_BASE");
    const newStations = vehiclesForNewStations === null
      ? null : Math.ceil(vehiclesForNewStations / vehiclesPerStation);
    const accessStatus = !accessVerified
      ? "Neověřeno" : accessPasses ? "Splňuje limit 1 000 m" : "Nesplňuje limit 1 000 m";
    return estimate([
      { label: "Obyvatelé řešeného území po záměru", value: projectPopulation, unit: "obyvatel" },
      { label: "Obyvatelé včetně spádu", value: catchmentPopulation, unit: "obyvatel" },
      { label: "Typ území", value: territory, unit: "" },
      { label: "Ověřená pěší vzdálenost k existující zbrojnici", value: accessVerified ? walkDistance : null, unit: "m" },
      { label: "Standard MMR – nejvýše skutečná pěší trasa", value: 1000, unit: "m" },
      { label: "Pěší dostupnost existující zbrojnice", value: accessStatus, unit: "" },
      { label: "Obyvatelé pokrytí stávající jednotkou JSDH", value: coveredIsValid ? covered : null, unit: "obyvatel" },
      { label: "Obyvatelé bez doloženého pokrytí", value: uncovered, unit: "obyvatel" },
      { label: "Odhadovaný počet vozidel potřebný pro spád",
        value: requiredVehicles === null ? null : `${formatNumber(requiredVehicles)} ${czechCountUnit(requiredVehicles, "vozidlo", "vozidla", "vozidel")}`, unit: "" },
      { label: "Návrhový normativ vozidel (scénář)", value: populationPerVehicle, unit: "obyvatel/vozidlo" },
      { label: "Doplňkový počet vozidel k zajištění",
        value: additionalVehicles === null ? null : `${formatNumber(additionalVehicles)} ${czechCountUnit(additionalVehicles, "vozidlo", "vozidla", "vozidel")}`, unit: "" },
      { label: "Modelová potřeba nových zbrojnic po kontrole pěší dostupnosti",
        value: newStations,
        unit: newStations === null ? "" : czechCountUnit(newStations, "zbrojnice JSDH", "zbrojnice JSDH", "zbrojnic JSDH") },
      { label: "Výsledek posouzení", value: needsReview
        ? "Potřeba JSDH zůstává započtena" : "Pěší dostupnost i kapacita vyhovují; nová zbrojnice se nezapočítá", unit: "" }
    ], [
      `Počet obyvatel ve spádu po záměru = ${catchmentPopulation === null ? "chybí vstup" : `${formatNumber(catchmentPopulation)} obyvatel`}; typ ${territory}.`,
      accessVerified
        ? `Ověřená skutečná pěší trasa ${formatNumber(walkDistance)} m ${accessPasses ? "splňuje" : "nesplňuje"} doporučený limit 1 000 m podle metodiky TAČR/MMR.`
        : "Skutečná pěší trasa k existující zbrojnici není doložena; do ověření zůstává potřeba započtena.",
      vehicleCapacityVerified
        ? `Scénářová poptávka = ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${requiredVehicles}; po odečtu ${availableVehicles} ověřených vozidel zbývá ${additionalVehicles}.`
        : `Bez doložené pěší dostupnosti do 1 000 m a ověřené volné kapacity se započítá celý scénář: ceil(${formatNumber(uncovered)} obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${requiredVehicles} vozidel.`,
      needsReview
        ? "Pěší standard popisuje dostupnost zbrojnice pro obyvatele; neprokazuje kategorii JPO, akceschopnost ani čas zásahu."
        : "Pěší dostupnost i zadaná kapacita vozidel vyhovují modelové potřebě; nová zbrojnice se v této bilanci nezapočítá."
    ], [
      "TAČR/MMR TB050MMR01, tabulka 1, tištěná str. 12: skutečná pěší docházka do 1 000 m k hasičské zbrojnici dobrovolných hasičů. Jde o nezávazný standard dostupnosti; nenahrazuje klasifikaci JPO, akceschopnost ani zásahový dojezd.",
      "Model kapacity, počtu vozidel a nové zbrojnice je scénářový. Plošné předpoklady a cena zbrojnice jsou nastaveny na polovinu výchozího modelu HZS a zůstávají upravitelné.",
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], vehiclesForNewStations, [uncovered === null
      ? "Chybí počet obyvatel bez doloženého pokrytí; požadovaný počet vozidel nelze vypočítat."
      : `Modelová potřeba vozidel = ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${formatNumber(requiredVehicles)} vozidel.`]);
  }

  if (item.id === "VYB_16A") {
    const covered = getTargetedPlanInput(item, "HZS_COVERED_POPULATION");
    const responseMinutes = getTargetedPlanInput(item, "HZS_RESPONSE_MINUTES");
    const availableVehicles = getTargetedPlanInput(item, "HZS_AVAILABLE_VEHICLES");
    const populationPerVehicle = getTargetedPlanInput(item, "MODEL_POPULATION_PER_VEHICLE");
    const source = getTargetedPlanSource(item, "HZS_COVERAGE_SOURCE");
    const responseLimit = ["A", "B"].includes(territory) ? 7 : 15;
    const validResponse = responseMinutes !== null;
    const coveredIsValid = covered !== null && catchmentPopulation !== null && covered <= catchmentPopulation;
    const uncovered = coveredIsValid ? Math.max(0, catchmentPopulation - covered) : catchmentPopulation;
    const requiredVehicles = uncovered === null ? null : Math.ceil(uncovered / populationPerVehicle);
    const vehicleCapacityVerified = availableVehicles !== null && Number.isInteger(availableVehicles)
      && availableVehicles >= 0 && validResponse && responseMinutes <= responseLimit;
    const additionalVehicles = requiredVehicles === null ? null : vehicleCapacityVerified
      ? Math.max(0, requiredVehicles - availableVehicles) : requiredVehicles;
    const needsReview = !validResponse || responseMinutes > responseLimit
      || !vehicleCapacityVerified || additionalVehicles > 0;
    const vehiclesForNewStations = requiredVehicles === null ? null : needsReview
      ? Math.max(1, additionalVehicles ?? 0) : 0;
    const newStations = vehiclesForNewStations === null ? null
      : Math.ceil(vehiclesForNewStations / getTargetedPlanInput(item, "MODEL_VEHICLES_PER_BASE"));
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
      { label: "Doplňkový počet vozidel k zajištění",
        value: additionalVehicles === null ? null : `${formatNumber(additionalVehicles)} ${czechCountUnit(additionalVehicles, "vozidlo", "vozidla", "vozidel")}`, unit: "" },
      { label: "Modelová potřeba nových stanic po kontrole dojezdu",
        value: newStations,
        unit: newStations === null ? "" : czechCountUnit(newStations, "stanice HZS", "stanice HZS", "stanic HZS") },
      { label: "Výsledek posouzení", value: needsReview ? "Potřeba HZS zůstává započtena" : "Dojezd vyhovuje; potřeba se nezapočítá", unit: "" }
    ], [
      `Počet obyvatel ve spádu po záměru = ${catchmentPopulation === null ? "chybí vstup" : `${formatNumber(catchmentPopulation)} obyvatel`}; typ ${territory}; plocha ${context.areaHa === null ? "chybí vstup" : `${formatNumber(context.areaHa)} ha`}.`,
      validResponse
        ? `Ověřený dojezd ${formatNumber(responseMinutes)} min ${responseMinutes <= responseLimit ? "splňuje" : "nesplňuje"} horní mez ${responseLimit} min pro typ ${territory}.`
        : `Dojezd není doložen; pro typ ${territory} se používá horní mez ${responseLimit} min a potřeba zůstává započtena do scénáře.`,
      vehicleCapacityVerified
        ? `Scénářová poptávka = ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${requiredVehicles}; po odečtu ${availableVehicles} ověřených vozidel zbývá ${additionalVehicles}.`
        : `Při absenci ověřeného volného vozidla se započítá celý scénář: ceil(${formatNumber(uncovered)} obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${requiredVehicles} vozidel.`,
      needsReview
        ? "Při chybějícím či nevyhovujícím dojezdu, nebo při doloženém nedostatku vozidel, zůstává potřeba započtena do posouzení."
        : "Doložený dojezd vyhovuje a zadaná kapacita vozidel nepřekračuje potřebu; nová stanice se v této bilanci nezapočítá."
    ], [
      "Metodika MMR/TB050MMR01, s. 12, odkazuje u hasičské stanice na dojezd podle přílohy k zákonu č. 133/1985 Sb.; pracovní list Cluster_Doprava_a_Bezpeci uvádí orientačně A/B 5–7 min a C/D 10–15 min. Poměr obyvatel na vozidlo, počet vozidel na stanici, plochy a cena jsou editovatelné scénáře, nikoli normy. Rozhoduje požární poplachový plán a HZS.",
      ...(source ? [`Zadaný podklad: ${source}.`] : [])
    ], vehiclesForNewStations, [uncovered === null
      ? "Chybí počet obyvatel bez doloženého pokrytí; požadovaný počet vozidel nelze vypočítat."
      : `Modelová potřeba vozidel = ceil(${formatNumber(uncovered)} nepokrytých obyvatel ÷ ${formatNumber(populationPerVehicle)} obyvatel/vozidlo) = ${formatNumber(requiredVehicles)} vozidel.`]);
  }

  if (item.id === "VYB_17A") {
    const responseMinutes = getTargetedPlanInput(item, "POLICE_RESPONSE_MINUTES");
    const source = getTargetedPlanSource(item, "POLICE_COVERAGE_SOURCE");
    const responseValid = ["A", "B", "C"].includes(territory) && responseMinutes !== null;
    const responseLimit = ["A", "B"].includes(territory) ? 10 : territory === "C" ? 15 : null;
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
      { label: "Modelová potřeba nových služeben", value: formatStations(proposedStations), unit: "" },
      { label: "Doporučená mez zásahu podle typu", value: responseLimit === null ? "D – neměří se" : `${formatNumber(responseLimit)} min`, unit: "" },
      { label: "Ověřená doba zásahu hlídky", value: responseValid ? responseMinutes : null, unit: "min" }
    ], [
      `Modelový počet policistů = ${catchmentPopulation === null ? "chybí počet obyvatel" : `${formatNumber(catchmentPopulation)} obyvatel ÷ ${formatNumber(populationPerOfficer)} obyvatel/policistu = ${formatNumber(modeledOfficers)} policistů`}; modelová služebna počítá nejvýše s ${formatNumber(officersPerStation)} policisty.`,
      `Typ ${territory}, plocha ${context.areaHa === null ? "chybí vstup" : `${formatNumber(context.areaHa)} ha`}, obyvatelé ve spádu ${catchmentPopulation === null ? "chybí vstup" : formatNumber(catchmentPopulation)}.`,
      responseLimit === null
        ? "Typ D nemá v tomto doporučení požadavek na dobu zásahu; modelovaný počet policistů a služeben zůstává zobrazen jako informativní scénář."
        : responseValid
          ? `Ověřený zásah hlídky ${formatNumber(responseMinutes)} min ${responseMinutes <= responseLimit ? "splňuje" : "nesplňuje"} doporučenou mez ${responseLimit} min pro typ ${territory}.`
          : `Doba zásahu hlídky pro typ ${territory} není doložena; modelovaný počet služeben zůstává ve scénáři.`,
      proposedStations === 0
        ? `Vyhovující ověřený zásah v typu ${territory} znamená nulovou modelovou potřebu nové služebny; personální potřeba se přesto uvádí pro kontrolu pokrytí.`
        : `Pro návrhovou výměru se uvažuje ${formatNumber(proposedStations)} služeben, každá pro nejvýše ${formatNumber(officersPerStation)} policistů; nejde o závazný služební normativ.`
    ], [
      "Doporučení interních metodik policie: typ A/B – zásah do 10 minut; typ C – do 15 minut; typ D – doba se neměří. Nejde o standard metodiky MMR ani o zákonnou povinnost zřídit služebnu. Zákon č. 273/2008 Sb. nestanoví kvótu policistů ani služeben na obyvatele; poměr obyvatel na policistu a kapacita služebny jsou scénářové předpoklady."
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
    const valid = demand !== null
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
  if (item.id === "VYB_13B") {
    const plan = getSupermarketEstimate(item);
    return plan.quantity === null
      ? "nelze vypočítat"
      : `${formatNumber(plan.quantity)} ${czechCountUnit(plan.quantity,
        "modelová prodejna", "modelové prodejny", "modelových prodejen")}`;
  }
  if (isHealthcareDemandItem(item)) {
    const isSpecialistGroup = item.id === "VYB_09B_1" || item.id === "VYB_09B_2";
    const label = item.id === "VYB_09A_4"
      ? "Poptávka po gynekologii po odečtu ověřené kapacity"
      : "Poptávka nepokrytá ověřenou kapacitou";
    const remaining = metricValue(label);
    const clinics = metrics.find((metric) => metric.label.startsWith("Modelový počet")
      && (isSpecialistGroup ? metric.label.includes("ambulanc")
        : metric.label.includes("ordinac")))?.value;
    if (clinics !== null && clinics !== undefined) {
      return `${formatNumber(clinics)} ${isSpecialistGroup
        ? czechCountUnit(clinics, "ambulance", "ambulance", "ambulancí")
        : czechCountUnit(clinics, "ordinace", "ordinace", "ordinací")}`;
    }
    return remaining === null || remaining === undefined
      ? "není vypočteno"
      : `${formatNumber(remaining, 2)} ${isSpecialistGroup ? "ambulancí" : "lékařských úvazků"}`;
  }
  if (["VYB_02C_1", "VYB_02C_2"].includes(item.id)) {
    const newCapacity = metricValue("Nová kapacita k zajištění");
    if (newCapacity !== null && newCapacity !== undefined)
      return `${formatNumber(newCapacity)} míst`;
    const unmetSeats = metricValue("Poptávka po odečtu ověřené volné kapacity");
    const verifiedFreeCapacity = metricValue("Ověřená volná kapacita") !== null
      && metricValue("Ověřená volná kapacita") !== undefined;
    const demand = metricValue("Odhadovaná poptávka po místech");
    if (unmetSeats !== null && unmetSeats !== undefined)
      return `${formatNumber(unmetSeats)} ${verifiedFreeCapacity ? "míst k doplnění" : "míst poptávky"}`;
    if (demand !== null && demand !== undefined)
      return `${formatNumber(demand)} žáků (scénář)`;
  }
  if (item.id === "VYB_09D") {
    const bases = metricValue("Modelová potřeba nových základen po kontrole dojezdu");
    return bases === null || bases === undefined
      ? "nelze vypočítat"
      : `${formatNumber(bases)} ${czechCountUnit(bases, "základna", "základny", "základen")}`;
  }
  if (item.id === "VYB_16A") {
    const stations = metricValue("Modelová potřeba nových stanic po kontrole dojezdu");
    return stations === null || stations === undefined
      ? "nelze vypočítat"
      : `${formatNumber(stations)} ${czechCountUnit(stations, "stanice", "stanice", "stanic")}`;
  }
  if (item.id === VOLUNTEER_FIRE_STATION_ID) {
    const stations = metricValue("Modelová potřeba nových zbrojnic po kontrole pěší dostupnosti");
    return stations === null || stations === undefined
      ? "nelze vypočítat"
      : `${formatNumber(stations)} ${czechCountUnit(stations, "zbrojnice JSDH", "zbrojnice JSDH", "zbrojnic JSDH")}`;
  }
  const status = metricValue("Výsledek podmíněného posouzení")
    ?? metricValue("Výsledek posouzení")
    ?? metricValue("Přítomnost v obci podle typu a populace");
  if (status !== null && status !== undefined) return String(status);
  if (item.id === "VYB_01_J") {
    const children = metrics.find((metric) => metric.label.startsWith("Odhad poptávky po péči"))?.value;
    return children === null || children === undefined
      ? "není vypočteno"
      : `${formatNumber(children, 1)} dětí (indikativně)`;
  }
  if (item.id === "VYB_09A_4") {
    const fte = metricValue("Poptávka po gynekologii po odečtu ověřené kapacity")
      ?? metricValue("Indikativní personální ekvivalent");
    return fte === null || fte === undefined
      ? "není vypočteno"
      : `${formatNumber(fte, 2)} lékařských úvazků`;
  }
  if (item.id === "VYB_17A") {
    const stations = metricValue("Modelová potřeba nových služeben");
    return stations === null || stations === undefined
      ? "není vypočteno"
      : String(stations);
  }
  if (item.id === "VYB_17B") {
    const deficit = metricValue("Poptávka nepokrytá ověřenou kapacitou úkrytů");
    return deficit === null || deficit === undefined
      ? "není vypočteno"
      : `${formatNumber(deficit)} osob`;
  }
  return "—";
}

function getSupermarketEstimate(item) {
  const localPopulation = numericOrNull(derivedValues().O_uzemi_cil);
  const projectIncrement = numericOrNull(getPrimary("O_celk"));
  const minimumPopulation = Math.max(0,
    numericOrNull(getItemEstimateValue(item, "SUPERMARKET_MIN_POPULATION")) ?? 3000);
  const populationPerStore = Math.max(1,
    numericOrNull(getItemEstimateValue(item, "SUPERMARKET_POPULATION_PER_STORE")) ?? 3000);
  const capacityVerified = getCapacityEdit(item).satisfiesNearbyCapacity === true;
  const totalPopulationMeetsThreshold = localPopulation !== null
    && localPopulation >= minimumPopulation;
  const quantity = localPopulation === null || projectIncrement === null
    ? null
    : !totalPopulationMeetsThreshold || projectIncrement <= 0 || capacityVerified
      ? 0
      : Math.max(1, Math.ceil(projectIncrement / populationPerStore));
  const usableAreaPerUnit = numericOrNull(getItemEstimateValue(item, "ITEM_USABLE_AREA_PER_UNIT"));
  const totalFloorFactor = numericOrNull(getItemEstimateValue(item, "ITEM_TOTAL_FLOOR_FACTOR"));
  const volumePerArea = numericOrNull(getItemEstimateValue(item, "ITEM_VOLUME_PER_AREA"));
  const siteAreaPerUnit = numericOrNull(getItemEstimateValue(item, "ITEM_SITE_AREA_PER_UNIT"));
  const pavedAreaPerUnit = numericOrNull(getItemEstimateValue(item, "ITEM_PAVED_AREA_PER_UNIT"));
  const rate = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_RATE"));
  const rateBasis = getItemEstimateValue(item, "ITEM_RTS_RATE_BASIS");
  const fixedCost = numericOrNull(getItemEstimateValue(item, "ITEM_RTS_FIXED_COST"));
  const surcharge = numericOrNull(getItemEstimateValue(item, "ITEM_COST_SURCHARGE"));
  const usableArea = quantity === null || usableAreaPerUnit === null
    ? null : quantity * usableAreaPerUnit;
  const totalFloorArea = usableArea === null || totalFloorFactor === null
    ? null : usableArea * totalFloorFactor;
  const volume = totalFloorArea === null || volumePerArea === null
    ? null : Math.round(totalFloorArea * volumePerArea);
  const siteArea = quantity === null || siteAreaPerUnit === null
    ? null : Math.round(quantity * siteAreaPerUnit);
  const pavedArea = quantity === null || pavedAreaPerUnit === null
    ? null : Math.round(quantity * pavedAreaPerUnit);
  const costBasis = rateBasis === "volume" ? volume : quantity;
  const zrn = rate === null || costBasis === null || fixedCost === null
    ? null : Math.round(costBasis * rate + fixedCost);
  const rtsRate = getRtsConstructionRate(item);
  const rateSource = Object.hasOwn(state.itemEstimateValues[item.key] || {}, "ITEM_RTS_RATE")
    ? "user-estimate"
    : rtsRate !== undefined
    ? rtsRate === null ? "missing" : "RTS"
    : item.rtsPricing?.rate ? "RTS" : "scenario-estimate";
  return {
    type: "supermarket",
    quantity,
    localPopulation,
    projectIncrement,
    minimumPopulation,
    populationPerStore,
    totalPopulationMeetsThreshold,
    capacityVerified,
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
    rateBasis,
    costBasis,
    fixedCost,
    surcharge,
    zrn,
    rateSource,
    cost: zrn === null || surcharge === null ? null : Math.round(zrn * (1 + surcharge)),
    unit: "prodejen"
  };
}

function getProjectPlanningEstimate(item, kind) {
  if (isInfrastructureReserveItem(item) && kind === "infrastructure") {
    const source = getInfrastructureReserveSource(item);
    if (!source) return null;
    const base = getProjectPlanningEstimate(source, "infrastructure");
    const percent = getInfrastructureReservePercent(item);
    if (!base || base.type !== "infrastructure" || base.quantity === null
      || base.quantity === undefined) return null;
    return {
      type: "infrastructure-reserve",
      sourceName: displayItemName(source),
      sourceQuantity: base.quantity,
      sourceCost: base.cost,
      sourceZrn: base.zrn,
      percent,
      quantity: base.quantity * percent,
      unit: base.unit,
      zrn: base.zrn === null || base.zrn === undefined
        ? null : Math.round(base.zrn * percent),
      cost: base.cost === null || base.cost === undefined
        ? null : Math.round(base.cost * percent)
    };
  }
  if (item.id === "VYB_13B") return getSupermarketEstimate(item);
  if (PARKING_SPLIT_ITEM_IDS.has(item.id) && ["legal", "recommended"].includes(kind)) {
    const calc = adjustedResult(item, kind);
    const quantity = numericOrNull(getSummaryRequirementQuantity(calc));
    if (quantity === null) return null;
    const share = Math.min(1, Math.max(0,
      numericOrNull(getParkingEstimateValue(item, kind, "PARKING_BUILDING_SHARE")) ?? 0));
    const volumePerStall = Math.max(0,
      numericOrNull(getParkingEstimateValue(item, kind, "PARKING_VOLUME_PER_STALL_M3")) ?? 0);
    const surfaceRatePerM3 = Math.max(0,
      numericOrNull(getParkingEstimateValue(item, kind, "PARKING_SURFACE_RATE_M3")) ?? 0);
    const pavedAreaPerStall = Math.max(0,
      numericOrNull(getParkingEstimateValue(item, kind, "PARKING_SURFACE_AREA_PER_STALL_M2")) ?? 0);
    const buildingStalls = quantity * share;
    const surfaceStalls = quantity - buildingStalls;
    const buildingVolume = buildingStalls * volumePerStall;
    const surfaceVolume = surfaceStalls * volumePerStall;
    const pavedArea = surfaceStalls * pavedAreaPerStall;
    const buildingRate = numericOrNull(getParkingEstimateValue(item, kind, "PARKING_BUILDING_RATE_M3"));
    const buildingCost = buildingRate === null ? null : Math.round(buildingVolume * buildingRate);
    const surfaceCost = Math.round(surfaceVolume * surfaceRatePerM3);
    const surcharge = Math.max(0,
      numericOrNull(getParkingEstimateValue(item, kind, "ITEM_COST_SURCHARGE")) ?? 0.08);
    const zrn = buildingCost === null ? null : buildingCost + surfaceCost;
    const cost = zrn === null ? null : Math.round(zrn * (1 + surcharge));
    return {
      type: "parking",
      quantity,
      buildingShare: share,
      buildingStalls,
      surfaceStalls,
      volumePerStall,
      buildingVolume,
      surfaceVolume,
      buildingRate,
      buildingCost,
      surfaceRatePerM3,
      surfaceCost,
      pavedAreaPerStall,
      pavedArea,
      siteArea: quantity * 28,
      zrn,
      costSurcharge: surcharge,
      cost,
      unit: "stání"
    };
  }
  if (isResidentParkingItem(item)) return null;
  if (item.id === "VYB_20" && kind === "recommended") {
    const healthcareItemIds = new Set(["VYB_09A_4", "VYB_09B_1", "VYB_09B_2", "VYB_17", "VYB_18", "VYB_19"]);
    const healthcareItems = [...healthcareItemIds].flatMap((id) => {
      const candidates = state.items.filter((clinic) => clinic.id === id);
      const recommended = candidates.find((clinic) => clinic.recommendedRules.length > 0);
      if (recommended) return [{ item: recommended, variant: "recommended" }];
      const legal = candidates.find((clinic) => clinic.legalRules.some((rule) =>
        !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id]))));
      return legal ? [{ item: legal, variant: "legal" }] : [];
    });
    const clinicEstimates = healthcareItems.map(({ item: clinic, variant }) => ({
      item: clinic,
      variant,
      estimate: getTargetedPlanEstimate(clinic, variant === "recommended")?.facilityPlan ?? null
    }));
    const discountCoefficient = numericOrNull(getProjectEstimateValue(item, "POLY_DISCOUNT_COEFFICIENT"));
    if (clinicEstimates.length === 0 || clinicEstimates.some((entry) =>
      !entry.estimate || !Number.isFinite(entry.estimate.cost)) || discountCoefficient === null) return null;
    const grossCost = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.cost, 0);
    const cost = Math.round(grossCost * discountCoefficient);
    const grossUsableArea = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.usableArea, 0);
    const grossTotalFloorArea = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.totalFloorArea, 0);
    const grossVolume = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.volume, 0);
    const grossSiteArea = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.siteArea, 0);
    const grossFootprintArea = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.footprintArea, 0);
    const grossPavedArea = clinicEstimates.reduce((sum, entry) => sum + entry.estimate.pavedArea, 0);
    return {
      type: "polyclinic",
      clinicEstimates,
      clinicCount: clinicEstimates.reduce((sum, entry) => sum + entry.estimate.count, 0),
      grossCost,
      discountCoefficient,
      discountSavings: Math.round(grossCost * (1 - discountCoefficient)),
      cost,
      grossUsableArea,
      grossTotalFloorArea,
      grossVolume,
      grossSiteArea,
      grossFootprintArea,
      grossPavedArea,
      usableArea: grossUsableArea * discountCoefficient,
      totalFloorArea: grossTotalFloorArea * discountCoefficient,
      volume: grossVolume * discountCoefficient,
      siteArea: grossSiteArea * discountCoefficient,
      footprintArea: grossFootprintArea * discountCoefficient,
      pavedArea: grossPavedArea * discountCoefficient
    };
  }
  if ((Object.hasOwn(TARGETED_PLAN_INPUTS, item.id)
      && (item.id !== "VYB_19" || isHealthcareItem(item)))
    || ["VYB_02C_1", "VYB_02C_2", "VYB_09A_4"].includes(item.id)) {
    return getTargetedPlanEstimate(item, kind === "recommended");
  }
  if (item.id === "VYB_04A" && ["legal", "recommended"].includes(kind)) {
    const qualifyingAreaHa = Math.max(0,
      numericOrNull(getItemEstimateValue(item, "PUBLIC_SPACE_QUALIFIED_AREA_HA"))
      ?? numericOrNull(getPrimary("S_zastav")) ?? 0);
    const pavedShare = Math.min(1, Math.max(0,
      numericOrNull(getItemEstimateValue(item, "PUBLIC_SPACE_PAVED_SHARE")) ?? 1));
    const pavingRate = Math.max(0,
      numericOrNull(getItemEstimateValue(item, "PUBLIC_SPACE_PAVING_RATE_M2")) ?? 0);
    const costSurcharge = Math.max(0,
      numericOrNull(getItemEstimateValue(item, "ITEM_COST_SURCHARGE")) ?? 0.08);
    const minimumArea = Math.ceil(qualifyingAreaHa * 500);
    const pavedArea = minimumArea * pavedShare;
    const zrn = Math.round(pavedArea * pavingRate);
    return {
      type: "public-space",
      qualifyingAreaHa,
      minimumArea,
      pavedShare,
      pavedArea,
      pavingRate,
      zrn,
      costSurcharge,
      cost: Math.round(zrn * (1 + costSurcharge)),
      unit: "m²"
    };
  }
  if (item.id === "VYB_04B" && kind === "recommended") {
    const parkCount = getSpatialPointRequirement(item, kind) ?? item.modelProposed ?? 1;
    const areaPerParkHa = Math.max(0.5,
      numericOrNull(getItemEstimateValue(item, "PARK_GREEN_AREA_HA")) ?? 0.5);
    const landscapingRate = Math.max(0,
      numericOrNull(getItemEstimateValue(item, "PARK_LANDSCAPING_RATE_M2")) ?? 500);
    const costSurcharge = Math.max(0,
      numericOrNull(getItemEstimateValue(item, "ITEM_COST_SURCHARGE")) ?? 0.08);
    const minimumArea = Math.ceil(parkCount * areaPerParkHa * 10000);
    const zrn = Math.round(minimumArea * landscapingRate);
    return {
      type: "public-space",
      isPark: true,
      parkCount,
      areaPerParkHa,
      minimumArea,
      pavedArea: 0,
      pavingRate: landscapingRate,
      zrn,
      costSurcharge,
      cost: Math.round(zrn * (1 + costSurcharge)),
      unit: "m²"
    };
  }
  if (item.id === HEATING_BOILER_ID && ["infrastructure", "review"].includes(kind)) {
    const energy = getAggregateProjectEnergyMetrics(item);
    const heatingShare = numericOrNull(getItemEstimateValue(item, "BOILER_HEATING_SHARE"));
    const fullLoadHours = numericOrNull(getItemEstimateValue(item, "BOILER_FULL_LOAD_HOURS"));
    const reserveFactor = numericOrNull(getItemEstimateValue(item, "BOILER_DESIGN_RESERVE"));
    const unitCapacity = numericOrNull(getItemEstimateValue(item, "BOILER_UNIT_CAPACITY_KW"));
    const heatDensity = numericOrNull(getItemEstimateValue(item, "BOILER_HEAT_DENSITY_KWH_M3_Y"));
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
      energyDensity: energy.energyDensity,
      peakToAverageRatio: numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO")),
      fallbackFloorArea: energy.fallbackFloorArea,
      fallbackBuildingVolume: energy.fallbackBuildingVolume,
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
    const population = Math.max(0, numericOrNull(derivedValues().O_uzemi_cil) ?? 0);
    const rainfallDepth = numericOrNull(getItemEstimateValue(item, "ALT_RAINFALL_DEPTH_MM"));
    const runoffCoefficient = numericOrNull(getItemEstimateValue(item, "ALT_RUNOFF_COEFFICIENT"));
    const roofRunoffCoefficient = numericOrNull(getItemEstimateValue(item, "ALT_ROOF_RUNOFF_COEFFICIENT"))
      ?? runoffCoefficient;
    const pavedRunoffCoefficient = numericOrNull(getItemEstimateValue(item, "ALT_PAVED_RUNOFF_COEFFICIENT"))
      ?? runoffCoefficient;
    const waterUse = numericOrNull(getItemEstimateValue(item, "ALT_WATER_LITERS_PER_PERSON_DAY"));
    const facilityWaterUse = numericOrNull(getItemEstimateValue(item, "ALT_FACILITY_WATER_LITERS_PER_M3_DAY"));
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
    const energyDensity = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_DENSITY_W_M3"));
    const peakToAverageRatio = numericOrNull(getItemEstimateValue(item, "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO"));
    const powerCableRate = numericOrNull(getItemEstimateValue(item, "ALT_POWER_CABLE_RATE"));
    const transformerCapacity = numericOrNull(getItemEstimateValue(item, "ALT_TRANSFORMER_CAPACITY_KW"));
    const transformerRate = numericOrNull(getItemEstimateValue(item, "ALT_TRANSFORMER_RATE"));
    const surcharge = numericOrNull(getItemEstimateValue(item, "ALT_COST_SURCHARGE"));
    if ([rainfallDepth, runoffCoefficient, roofRunoffCoefficient, pavedRunoffCoefficient,
      waterUse, facilityWaterUse, peakFactor, wastewaterReturn,
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
    let lineSizingBase = 0;
    let lineSizingRatio = 0;
    let lineSizingExponent = 0;
    const roofRunoffArea = footprint.roofArea * roofRunoffCoefficient;
    const pavedRunoffArea = footprint.pavedArea * pavedRunoffCoefficient;
    const effectiveRunoffArea = roofRunoffArea + pavedRunoffArea;
    const facilityWaterLitersPerDay = footprint.totalFacilityVolume * facilityWaterUse;
    const averageDailyWater = (population * waterUse + facilityWaterLitersPerDay) / 1000;
    const averageDailyWastewater = averageDailyWater * wastewaterReturn;
    const sizeLineCost = (baselineCapacity, exponent) => {
      lineSizingBase = baselineCapacity;
      lineSizingExponent = exponent;
      lineSizingRatio = baselineCapacity > 0
        ? Math.max(0, quantity) / baselineCapacity
        : quantity > 0 ? 1 : 0;
      lineCost = networkLength * lineRate * lineSizingRatio ** exponent;
    };
    if (item.id === "VYB_20A") {
      metric = effectiveRunoffArea * rainfallDepth / 1000;
      quantity = metric;
      unit = "m³ retenčního objemu";
      lineRate = rainwaterRate;
      lineCost = quantity * lineRate;
      pointLabel = "Nádrž a stavební práce";
    } else if (item.id === "VYB_20B_1") {
      metric = averageDailyWater;
      quantity = metric * peakFactor;
      unit = "m³/den návrhové kapacity";
      lineRate = waterPipeRate;
      sizeLineCost(population * waterUse / 1000 * peakFactor, 0.5);
      pointCount = waterPointCount;
      pointRate = waterPointRate;
      pointLabel = "Vodoměrná a armaturní šachta";
    } else if (item.id === "VYB_20B_2") {
      metric = averageDailyWastewater;
      quantity = metric * peakFactor;
      unit = "m³/den návrhové kapacity";
      lineRate = sewerPipeRate;
      sizeLineCost(population * waterUse / 1000 * wastewaterReturn * peakFactor, 0.5);
      pointCount = networkLength > 0 ? Math.ceil(networkLength / sewerManholeSpacing) + 1 : 0;
      pointRate = sewerManholeRate;
      pointLabel = "Revizní šachty";
    } else {
      energyMetrics = getAggregateProjectEnergyMetrics(item);
      metric = energyMetrics.peakElectricalLoad;
      quantity = metric;
      unit = "kW návrhového příkonu";
      lineRate = powerCableRate;
      sizeLineCost(getAggregateProjectEnergyMetrics(item, false).peakElectricalLoad, 1);
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
      roofRunoffArea,
      pavedRunoffArea,
      effectiveRunoffArea,
      roofRunoffCoefficient,
      pavedRunoffCoefficient,
      totalFloorArea: footprint.totalFloorArea,
      totalFacilityVolume: footprint.totalFacilityVolume,
      facilityWaterLitersPerDay,
      facilityDailyWater: facilityWaterLitersPerDay / 1000,
      residentialDailyWater: population * waterUse / 1000,
      facilityWaterUse,
      averageDailyWater,
      averageDailyWastewater,
      annualElectricity: energyMetrics?.annualElectricity ?? null,
      modeledAnnualElectricity: energyMetrics?.modeledAnnualElectricity ?? null,
      modeledPeakElectricalLoad: energyMetrics?.modeledPeakElectricalLoad ?? null,
      modeledEquipmentCount: energyMetrics?.modeledEquipmentCount ?? 0,
      fallbackFloorArea: energyMetrics?.fallbackFloorArea ?? 0,
      fallbackBuildingVolume: energyMetrics?.fallbackBuildingVolume ?? 0,
      fallbackPeakElectricalLoad: energyMetrics?.fallbackPeakElectricalLoad ?? 0,
      peakToAverageRatio,
      sourceMetric: metric,
      quantity,
      unit,
      networkLength,
      lineRate,
      lineSizingBase,
      lineSizingRatio,
      lineSizingExponent,
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
      roofRunoffCoefficient,
      pavedRunoffCoefficient,
      waterUse,
      facilityWaterUse,
      peakFactor,
      wastewaterReturn,
      energyDensity,
      transformerCapacity
    };
  }

  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
  const baseUsableArea = numericOrNull(item.usableArea);
  if (baseUsableArea === null) return null;
  const recommendation = getMmrSchoolRecommendation(item);
  const targetCapacity = kind === "recommended"
    ? getPlannedQuantity(item, kind) : numericOrNull(item.demand);
  const modelReferenceDemand = recommendation?.modelReferenceDemand;
  if (targetCapacity === null || modelReferenceDemand === null || modelReferenceDemand <= 0) return null;
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
  const scale = targetCapacity / modelReferenceDemand;
  const standardUsableArea = getStandardMetricValue(baseUsableArea, scale);
  const upliftPercent = 0;
  const reserveUsableArea = 0;
  const usableArea = standardUsableArea;
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
    modelReferenceDemand,
    targetCapacity,
    recommendation: kind === "recommended" ? recommendation : null,
    scale,
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
  } else if (item.socialDemographic) {
    rateSource = estimateComparableUnitRate(item) !== null ? "comparable-estimate" : "missing";
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
      && (!HEALTHCARE_DEMAND_ITEM_IDS.has(item.id)
        || (kind === "legal"
          ? field.code !== "MODEL_POPULATION_PER_FTE"
            && field.code !== "MODEL_MINIMUM_POPULATION_PER_FTE"
            && field.code !== "MODEL_MINIMUM_DENSITY_PER_HA"
            && ![
              "GYNECOLOGY_WALK_DISTANCE", "GYNECOLOGY_DRIVE_MINUTES", "GYNECOLOGY_ACCESS_SOURCE",
              "HEALTHCARE_WALK_DISTANCE", "HEALTHCARE_DRIVE_MINUTES", "HEALTHCARE_ACCESS_SOURCE"
            ].includes(field.code)
          : field.code !== "MODEL_LEGAL_PATIENTS_PER_DOCTOR"))
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
    fields = PROJECT_ESTIMATE_INPUTS.filter((field) =>
      field.code === "POLY_DISCOUNT_COEFFICIENT");
    note = "Z každého druhu péče se započítá doporučená varianta, pokud existuje, jinak zákonná. Koeficient efektivity 0,7 je editovatelný a snižuje souhrnné plochy a investici; ordinace ani jejich kapacitu nesnižuje.";
  } else if (estimate?.type === "infrastructure") {
    fields = ALTERNATIVE_ESTIMATE_INPUTS.filter((field) => {
      if (item.id === HEATING_BOILER_ID) return [
        "ALT_ENERGY_DENSITY_W_M3", "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO",
        "BOILER_HEAT_DENSITY_KWH_M3_Y", "BOILER_HEATING_SHARE",
        "BOILER_FULL_LOAD_HOURS", "BOILER_DESIGN_RESERVE",
        "BOILER_UNIT_CAPACITY_KW", "BOILER_UNIT_COST", "ALT_COST_SURCHARGE"
      ].includes(field.code);
      if (item.id === "VYB_20A") return [
        "ALT_RAINFALL_DEPTH_MM", "ALT_ROOF_RUNOFF_COEFFICIENT",
        "ALT_PAVED_RUNOFF_COEFFICIENT", "ALT_RAINWATER_RATE", "ALT_COST_SURCHARGE"
      ].includes(field.code);
      if (item.id === "VYB_20B_1") return [
        "ALT_WATER_LITERS_PER_PERSON_DAY", "ALT_FACILITY_WATER_LITERS_PER_M3_DAY",
        "ALT_PEAK_DAY_FACTOR", "ALT_NETWORK_LENGTH_M",
        "ALT_WATER_PIPE_RATE", "ALT_WATER_METER_CHAMBER_COUNT", "ALT_WATER_METER_CHAMBER_RATE",
        "ALT_COST_SURCHARGE"
      ].includes(field.code);
      if (item.id === "VYB_20B_2") return [
        "ALT_WATER_LITERS_PER_PERSON_DAY", "ALT_FACILITY_WATER_LITERS_PER_M3_DAY",
        "ALT_PEAK_DAY_FACTOR", "ALT_WASTEWATER_RETURN",
        "ALT_NETWORK_LENGTH_M", "ALT_SEWER_PIPE_RATE", "ALT_SEWER_MANHOLE_SPACING_M",
        "ALT_SEWER_MANHOLE_RATE", "ALT_COST_SURCHARGE"
      ].includes(field.code);
      return [
        "ALT_ENERGY_DENSITY_W_M3", "ALT_ENERGY_PEAK_TO_AVERAGE_RATIO",
        "ALT_NETWORK_LENGTH_M", "ALT_POWER_CABLE_RATE",
        "ALT_TRANSFORMER_CAPACITY_KW", "ALT_TRANSFORMER_RATE", "ALT_COST_SURCHARGE"
      ].includes(field.code);
    });
    note = item.id === HEATING_BOILER_ID
      ? "Jde o scénářový odhad, nikoli projektový výpočet. Využívá roční spotřebu elektřiny a příkon modelovaných položek ze sešitu; chybějící údaje se odhadují podle objemu realizovaných budov. Cena modulu je orientačně odvozena sazbou 10 000 Kč/kW z jeho zadaného výkonu (výchozích 500 kW = 5 mil. Kč/modul); nejde o sazbu RTS a nezahrnuje projekt, kotelnu ani související technologie. Neelektrická energie není automaticky totožná s potřebou vytápění. Upravte předpoklady podle energetického posudku."
      : "Odhad mimo RTS a projektový rozpočet. Trasa je výchozím předpokladem 1,5 km. Sazby za metr jsou odvozeny z rozdělení nákladového scénáře TECH-COST-001 na tuto délku; šachty a trafostanice jsou samostatné pracovní odhady. Spotřeba domácností vychází z ČSÚ za rok 2023 (86,7, zaokrouhleno na 87 l/os./den). Všechny sazby a výměry upravte podle podmínek správců sítí a skutečné trasy.";
  } else if (item.id === "VYB_07B") {
    if (kind !== "recommended") return null;
    fields = HOUSING_CONTRIBUTION_INPUTS;
    note = "Editovatelné scénářové koeficienty; podíl městského odkupu a ušlá marže se nastavují v hlavních vstupech.";
  } else if (estimate?.type === "public-space") {
    fields = ITEM_ESTIMATE_INPUTS.filter((field) => estimate.isPark
      ? ["PARK_GREEN_AREA_HA", "PARK_LANDSCAPING_RATE_M2", "ITEM_COST_SURCHARGE"].includes(field.code)
      : ["PUBLIC_SPACE_QUALIFIED_AREA_HA", "PUBLIC_SPACE_PAVED_SHARE",
        "PUBLIC_SPACE_PAVING_RATE_M2", "ITEM_COST_SURCHARGE"].includes(field.code));
    note = estimate.isPark
      ? "Výchozí plocha jednoho parku je dolní hranice 0,5 ha z MMR; plocha i agregovaná cena parkové úpravy jsou editovatelné scénářové předpoklady mimo RTS."
      : "Základ plochy přebírá S_zastav, ale lze jej upravit. Podíl zpevnění a cena jsou editovatelné scénářové odhady; sazba není položkou RTS.";
  } else if (estimate?.type === "parking") {
    fields = PARKING_ESTIMATE_INPUTS;
    note = "Podíl parkovacího domu je editovatelné procento; zbytek stání se počítá jako venkovní parkování na zpevněné ploše. Výchozí hodnoty jsou 20 % u zákonné a 80 % u doporučené varianty. Sazba domu vychází z RTS 2026, JKSO 811.5; venkovní stání používá pracovní objemovou sazbu z nákladového modelu sešitu.";
  } else if (estimate?.type === "supermarket") {
    fields = SUPERMARKET_ESTIMATE_INPUTS;
    note = "Uživatelský scénář: výchozí práh 3 000 obyvatel a jedna prodejna na 3 000 nových obyvatel. Jde o předpoklad, nikoli standard MMR nebo zákonnou povinnost. Plocha a cena jsou editovatelné; cena využívá RTS JKSO 801.8, pokud je sazba pro zvolenou konstrukci dostupná.";
  } else if (estimate?.type === "item") {
    fields = ITEM_ESTIMATE_INPUTS.filter((field) =>
      (GENERIC_ITEM_ESTIMATE_CODES.has(field.code)
        || item.socialDemographic && field.code === "SOCIAL_DEMOGRAPHIC_COEFFICIENT")
      && !(item.socialDemographic && field.code === "ITEM_QUANTITY")
      && (!getRtsBuildingGroup(item)
        || !["ITEM_RTS_RATE", "ITEM_RTS_RATE_BASIS"].includes(field.code)));
    note = item.socialDemographic
      ? `Kapacita = počet obyvatel × demografický koeficient. Výměrové a cenové parametry zůstávají upravitelné níže. ${item.socialDemographic.source}`
      : estimate.rateSource === "RTS"
      ? "Předpoklady se ukládají pro tuto položku. Cena vychází z RTS; chybějící výměrové koeficienty jsou pracovní odhady."
      : estimate.rateSource === "missing"
        ? "Předpoklady se ukládají pro tuto položku. Cena není v RTS ani v porovnatelných datech; sazbu lze zadat jako vlastní odhad mimo RTS."
        : estimate.rateSource === "scenario-estimate"
            ? item.id === "VYB_04B"
              ? "RTS v dostupném podkladu cenu veřejné parkové zeleně neuvádí. Výchozí sazba 2 500 000 Kč/park (5 000 m²) je scénářový odhad 500 Kč/m² před přirážkou 8 % a odpovídá dolní hranici 0,5 ha z MMR; upravte ji podle návrhu a místních cen."
              : "Předpoklady se ukládají pro tuto položku. Výchozí cena i výměrové koeficienty jsou orientační odhady mimo RTS; upravte je podle záměru a místních cen."
        : "Předpoklady se ukládají pro tuto položku. Cena je orientační cenový odhad mimo RTS; ověřte ji podle nabídky nebo projektu.";
  }
  if (["item", "supermarket", "school"].includes(estimate?.type)) {
    note += " Koeficienty užitné plochy a obestavěného prostoru určují výměry a při cenovém základu Kč/m³ také cenu. Areál a zpevněná plocha jsou prostorové bilance bez samostatné sazby. Investice používá zvolený cenový základ, fixní náklad a přirážku.";
  }
  if (!fields.length) return null;

  const disclosure = createElement("details", "project-estimate-controls");
  disclosure.open = Boolean(item.socialDemographic);
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
    const value = estimate?.type === "parking"
      ? getParkingEstimateValue(item, kind, field.code)
      : ["item", "supermarket"].includes(estimate?.type) || estimate?.type === "infrastructure"
      ? getItemEstimateValue(item, field.code)
      : field.code.startsWith("K_") || field.code.startsWith("POLY_")
        ? getProjectEstimateValue(item, field.code)
        : getItemEstimateValue(item, field.code);
    const itemField = estimate?.type === "parking"
      ? { ...field, code: `${field.code}_${kind.toUpperCase()}` }
      : estimate?.type === "targeted"
      ? { ...field, color: TARGETED_FIELD_COLORS[field.code] || field.color }
      : item.socialDemographic && field.code === "SOCIAL_DEMOGRAPHIC_COEFFICIENT"
      ? {
        ...field,
        label: `Demografický koeficient kapacity (${item.socialDemographic.unitLabel} / obyvatele)`,
        min: 0,
        max: 1,
        step: 0.00001,
        unit: `${item.socialDemographic.unitLabel} / obyvatele`
      }
      : ["item", "supermarket", "school"].includes(estimate?.type) && field.code === "ITEM_RTS_RATE"
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
      `Nabídka obsahuje jen konstrukční charakteristiky s cenou v odpovídající skupině RTS 2026; zvolená sazba v Kč/m³ OP se automaticky promítne do výpočtu. Zdroj: RTS 2026, s. 1, JKSO ${getRtsBuildingGroup(item)}.`));
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
  if (estimate.type === "parking") {
    addMetric("Počet stání v propočtu", estimate.quantity, "stání");
    addMetric("Podíl parkovací dům / zpevněná plocha",
      `${formatNumber(estimate.buildingShare * 100)} / ${formatNumber((1 - estimate.buildingShare) * 100)}`, "%");
    addMetric("Stání v parkovacím domě", estimate.buildingStalls, "stání");
    addMetric("Stání na zpevněné ploše", estimate.surfaceStalls, "stání");
    addMetric("Obestavěný prostor parkovacího domu", estimate.buildingVolume, "m³");
    addMetric("Objemový ekvivalent venkovních stání", estimate.surfaceVolume, "m³");
    addMetric("Zpevněná plocha venkovního parkování", estimate.pavedArea, "m²");
    addMetric("Modelová investice před VRN a rezervou", estimate.zrn, "Kč");
    addMetric("Modelová investice včetně VRN a rezervy", estimate.cost, "Kč");
    return grid;
  }
  if (estimate.type === "infrastructure-reserve") {
    addMetric(`Rezerva kapacity ${estimate.sourceName}`,
      estimate.quantity, estimate.unit);
    addMetric("Procento rezervy", estimate.percent * 100, "%");
    addMetric("Orientační investice rezervy", estimate.cost, "Kč");
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
    addMetric("Zahrnuté druhy ordinací", estimate.clinicEstimates.length, "druhů");
    addMetric("Modelová kapacita ordinací a ambulancí celkem", estimate.clinicCount, "ordinací / ambulancí");
    for (const { item: clinic, variant, estimate: clinicEstimate } of estimate.clinicEstimates) {
      addMetric(`${displayItemName(clinic)} – ${variant === "recommended" ? "doporučeno" : "zákonný základ"}`,
        `${clinicEstimate.metrics[0].value}; ${clinicEstimate.metrics[1].value}`, "");
    }
    addMetric("Součet užitných ploch před koeficientem", estimate.grossUsableArea, "m²");
    addMetric("Užitná plocha polikliniky po koeficientu", estimate.usableArea, "m²");
    addMetric("Celková podlahová plocha po koeficientu", estimate.totalFloorArea, "m²");
    addMetric("Zastavěná plocha po koeficientu", estimate.footprintArea, "m²");
    addMetric("Obestavěný prostor po koeficientu", estimate.volume, "m³");
    addMetric("Plocha areálu po koeficientu", estimate.siteArea, "m²");
    addMetric("Zpevněná plocha po koeficientu", estimate.pavedArea, "m²");
    addMetric("Součet investic ordinací před koeficientem", estimate.grossCost, "Kč");
    addMetric("Koeficient efektivity společné polikliniky", estimate.discountCoefficient, "násobek");
    addMetric("Modelovaná úspora investice", estimate.discountSavings, "Kč");
    addMetric("Orientační investice polikliniky", estimate.cost, "Kč");
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
  const estimate = getTargetedPlanEstimate(item, kind === "recommended");
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
      : estimateType === "supermarket" ? SUPERMARKET_ESTIMATE_INPUTS
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
  const controls = estimate?.type === "infrastructure-reserve"
    ? null : estimate ? renderProjectEstimateControls(item, kind, estimate) : null;
  if (controls) section.append(controls);
  if (item.socialDemographic) {
    const objectAssumptions = renderObjectAssumptionControls(item, kind);
    if (objectAssumptions) section.append(objectAssumptions);
  }
  openingEquations.forEach((equation) => {
    if (item.socialDemographic) {
      section.append(createElement("p", "calculation-equation", equation));
    } else {
      appendCalculationText(section, equation, item, kind);
    }
  });
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
  if (estimate.type === "infrastructure-reserve") {
    addLine(`Základní kapacita · ${estimate.sourceName}`,
      [`${formatNumber(estimate.sourceQuantity)} ${text(estimate.unit, "")}`],
      `${formatNumber(estimate.sourceQuantity)} ${text(estimate.unit, "")}`, "volume");
    addLine("Kapacita navíc jako samostatná rezerva",
      [`${formatNumber(estimate.sourceQuantity)} × ${formatNumber(estimate.percent * 100)} %`],
      `${formatNumber(estimate.quantity)} ${text(estimate.unit, "")}`, "area");
    addLine("Orientační investice rezervy",
      estimate.sourceCost === null || estimate.sourceCost === undefined
        ? ["Cena základní infrastruktury není stanovena"]
        : [`${formatCost(estimate.sourceCost)} × ${formatNumber(estimate.percent * 100)} %`],
      formatCost(estimate.cost), "rts");
    section.append(createElement("p", "calculation-source-note",
      "Rezerva je samostatná scénářová položka: její kapacita i orientační cena jsou stejným procentem odvozeny od vypočtené základní infrastruktury. Při výběru rezervy se základní položka automaticky odebere z realizace."));
    return section;
  }
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

  if (estimate.type === "parking") {
    addLine("Počet stání v parkovacím domě",
      [`${formatNumber(estimate.quantity)} stání × ${formatNumber(estimate.buildingShare * 100)} %`],
      `${formatNumber(estimate.buildingStalls)} stání`, "site");
    addLine("Počet venkovních stání",
      [`${formatNumber(estimate.quantity)} − ${formatNumber(estimate.buildingStalls)} stání`],
      `${formatNumber(estimate.surfaceStalls)} stání`, "paved");
    addLine("Objem parkovacího domu",
      [`${formatNumber(estimate.buildingStalls)} stání × ${formatNumber(estimate.volumePerStall)} m³/stání`],
      `${formatNumber(estimate.buildingVolume)} m³`, "volume");
    if (estimate.buildingRate === null) {
      addMissingLine("Cena parkovacího domu", "pro zvolenou konstrukční charakteristiku není v RTS sazba; zadejte sazbu v editovatelných předpokladech.", "rts");
    } else {
      addLine("Cena parkovacího domu dle RTS",
        [`${formatNumber(estimate.buildingVolume)} m³ × ${formatCost(estimate.buildingRate)}/m³`],
        formatCost(estimate.buildingCost), "rts");
    }
    addLine("Cena venkovních stání",
      [`${formatNumber(estimate.surfaceVolume)} m³ × ${formatCost(estimate.surfaceRatePerM3)}/m³ (modelový ekvivalent)`],
      formatCost(estimate.surfaceCost), "rts");
    if (estimate.zrn !== null) {
      addLine("Základní náklady (ZRN)",
        [`${formatCost(estimate.buildingCost)} + ${formatCost(estimate.surfaceCost)}`],
        formatCost(estimate.zrn), "rts");
      addLine("Modelová investice včetně VRN a rezervy",
        [`${formatCost(estimate.zrn)} × (1 + ${formatNumber(estimate.costSurcharge * 100)} %)`],
        formatCost(estimate.cost), "surcharge");
    }
    sources.push(`${RTS_811_5_SOURCE} Konstrukční charakteristika ${getItemEstimateValue(item, JKSO_CONSTRUCTION_INPUT.code)}; sazba ${formatCost(estimate.buildingRate)}/m³.`);
    sources.push("Venkovní stání: editovatelná sazba 7 500 Kč/m³ z nákladového modelu pracovního sešitu; objem na stání je pouze modelový ekvivalent pro srovnání, nikoli výměra parkoviště ani rozpočet. Podíl je defaultně 20 % pro zákonné a 80 % pro doporučené parkování.");
    section.append(createElement("p", "calculation-source-note", `Zdroj a povaha odhadu: ${[...new Set(sources)].join(" ")}`));
    return section;
  }

  if (estimate.type === "public-space") {
    if (estimate.isPark) {
      addLine("Počet parků podle odhadu pěšího pokrytí",
        [`${formatNumber(estimate.parkCount)} bodů`],
        `${formatNumber(estimate.parkCount)} parků`, "site");
      addLine("Plocha parkové zeleně",
        [`${formatNumber(estimate.parkCount)} parků × ${formatNumber(estimate.areaPerParkHa)} ha/park`],
        `${formatNumber(estimate.minimumArea)} m²`, "site");
      addLine("Agregovaná parková úprava",
        [`${formatNumber(estimate.minimumArea)} m² × ${formatCost(estimate.pavingRate)}/m²`],
        formatCost(estimate.zrn), "rts");
      addLine("Orientační investice včetně VRN a rezervy",
        [`${formatCost(estimate.zrn)} × (1 + ${formatNumber(estimate.costSurcharge * 100)} %)`],
        formatCost(estimate.cost), "surcharge");
      sources.push("TB050MMR01, tabulka 4, tištěná s. 17: parková zeleň v rozsahu 0,5–1 ha; typy A/B/C do 300 m skutečné pěší vzdálenosti, typ D v každém obytném sídle. Počet bodů je geometrický screening, nikoli náhrada GIS.");
      sources.push("Sazba parkové úpravy je editovatelný agregovaný scénář mimo RTS, nikoli položkový rozpočet.");
      section.append(createElement("p", "calculation-source-note", `Zdroj a podmínka použití: ${[...new Set(sources)].join(" ")}`));
      return section;
    }
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
    addLine(kind === "recommended"
      ? "Doporučená výměra veřejného prostranství MMR"
      : "Zákonné minimum veřejného prostranství",
      [`${formatNumber(estimate.qualifyingAreaHa)} ha × 500 m²/ha`],
      `${formatNumber(estimate.minimumArea)} m²`, "area");
    addLine("Plocha oceněná jako zpevněná",
      [`${formatNumber(estimate.minimumArea)} m² × ${formatNumber(estimate.pavedShare * 100)} %`],
      `${formatNumber(estimate.pavedArea)} m²`, "paved");
    addLine("Modelová cena zpevněné plochy",
      [`${formatNumber(estimate.pavedArea)} m² × ${formatCost(estimate.pavingRate)}/m²`],
      formatCost(estimate.zrn), "rts");
    addLine("Orientační investice včetně VRN a rezervy",
      [`${formatCost(estimate.zrn)} × (1 + ${formatNumber(estimate.costSurcharge * 100)} %)`],
      formatCost(estimate.cost), "surcharge");
    sources.push(kind === "recommended"
      ? "TB050MMR01, tabulka 1 (tištěná str. 17): doporučení nejméně 1 000 m² veřejného prostranství na každé 2 ha související zastavitelné plochy; pozemní komunikace se nezapočítávají."
      : "§ 7 odst. 2 vyhlášky č. 501/2006 Sb.: nejméně 1 000 m² veřejného prostranství na každé 2 ha související zastavitelné plochy bydlení, rekreace, občanského vybavení nebo smíšené obytné. Pozemní komunikace se do výměry nezapočítávají.");
    sources.push("Plošná cena zpevnění je editovatelný pracovní odhad mimo RTS, nikoli položkový rozpočet; ověřte skladbu konstrukce, materiál a místní ceny.");
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
        [`${formatNumber(estimate.fallbackBuildingVolume)} m³ obestavěného prostoru × ${formatNumber(estimate.heatDensity)} kWh/m³/rok`],
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
      sources.push("Rozdíl celkové a elektrické energie je pouze modelový neelektrický energetický proxy, nikoli ověřená potřeba tepla. Pro položky bez energetických hodnot používá výpočet editovatelnou náhradní potřebu elektřiny a tepla na m³ obestavěného prostoru; zdrojová tabulka ani cenový podklad kotelny nejsou úplné.");
      sources.push(`Ekvivalentní hodiny plného zatížení, podíl vytápění, výkonová rezerva, modulová kapacita a cena ${formatNumber(estimate.pointRate)} Kč/modul (${formatNumber(estimate.pointRate / estimate.unitCapacity)} Kč/kW) jsou upravitelné scénářové předpoklady. Jednotková cena je odvozena z výkonu modulu sazbou 10 000 Kč/kW, není doložena RTS ani nabídkou dodavatele a nezahrnuje kompletní technologii či stavební část. Návrhový výkon a počet modulů je nutné nahradit výpočtem podle energetického posudku a projektu.`);
      section.append(createElement("p", "calculation-source-note", `Zdroj a omezení: ${[...new Set(sources)].join(" ")}`));
      return section;
    } else if (item.id === "VYB_20A") {
      addLine("Příspěvek střech k odtoku",
        [`${formatNumber(surface.roofArea)} m² × ${formatNumber(estimate.roofRunoffCoefficient)}`],
        `${formatNumber(estimate.roofRunoffArea)} m²`, "volume");
      addLine("Příspěvek zpevněných ploch k odtoku",
        [`${formatNumber(surface.pavedArea)} m² × ${formatNumber(estimate.pavedRunoffCoefficient)}`],
        `${formatNumber(estimate.pavedRunoffArea)} m²`, "paved");
      addLine("Ekvivalentní odtoková plocha",
        [`${formatNumber(estimate.roofRunoffArea)} m² střech + ${formatNumber(estimate.pavedRunoffArea)} m² zpevněných ploch`],
        `${formatNumber(estimate.effectiveRunoffArea)} m²`, "area");
      addLine("Záchytný objem srážkové vody",
        [`${formatNumber(estimate.effectiveRunoffArea)} m² × ${formatNumber(estimate.rainfallDepth)} mm ÷ 1 000`],
        `${formatNumber(estimate.quantity)} m³`, "area");
      sources.push("Modelová srážka 15 mm odpovídá pracovnímu údaji 150 m³/ha v listu TECH-HYDRO-001. Střechy používají výchozí součinitel 0,95 a zpevněné plochy nižší součinitel 0,65; oba jsou editovatelné odhady.");
      sources.push("Střešní plocha základního bytového fondu se odvozuje z HPP_byty a průměrné podlažnosti; další střechy a zpevněné plochy se přičítají z realizovaných vybaveností.");
      sources.push("Sazba 13 600 Kč/m³ je cenový proxy z TECH-COST-001, nikoli ověřená tržní cena konkrétní nádrže. Obsah výkopů, konstrukce a napojení ověřit projektem.");
    } else if (item.id === "VYB_20B_1" || item.id === "VYB_20B_2") {
      addLine("Průměrná denní potřeba vody obyvatel",
        [`${formatNumber(estimate.population)} obyvatel × ${formatNumber(estimate.waterUse)} l/osobu/den ÷ 1 000`],
        `${formatNumber(estimate.residentialDailyWater)} m³/den`, "area");
      addLine("Průměrná denní potřeba vody realizované vybavenosti",
        [`${formatNumber(estimate.totalFacilityVolume)} m³ obestavěného prostoru × ${formatNumber(estimate.facilityWaterUse)} l/m³/den ÷ 1 000`],
        `${formatNumber(estimate.facilityDailyWater)} m³/den`, "volume");
      addLine("Celková průměrná denní potřeba vody",
        [`${formatNumber(estimate.residentialDailyWater)} + ${formatNumber(estimate.facilityDailyWater)} m³/den`],
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
      sources.push("Obytná spotřeba se počítá z populace po záměru; byty prodávané městu jsou v ní zahrnuty vždy. Spotřeba vybaveností se přičítá podle objemu realizovaných budov a editovatelného objemového proxy 1,5 l/m³/den.");
      const lengthLabel = item.id === "VYB_20B_1" ? "Délka nového vodovodního potrubí" : "Délka nové splaškové kanalizace";
      const pipeLabel = item.id === "VYB_20B_1" ? "Cena vodovodního potrubí včetně pokládky" : "Cena splaškového potrubí včetně pokládky";
      addLine(lengthLabel, [], `${formatNumber(estimate.networkLength)} m`, "site");
      addLine(pipeLabel,
        [`${formatNumber(estimate.networkLength)} m × ${formatCost(estimate.lineRate)}/m × ${formatNumber(estimate.lineSizingRatio)}^${formatNumber(estimate.lineSizingExponent)} (poměr návrhové a rezidenční kapacity)`],
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
      sources.push("Aby se odhad ceny změnil s poptávkou, základní cenu trasy násobí odmocnina poměru celkové a rezidenční návrhové kapacity. Jde o transparentní proxy dimenze potrubí, nikoli výkaz výměr nebo ceníkovou nabídku.");
    } else {
      addLine("Souhrnná podlahová plocha objektů", [], `${formatNumber(surface.totalFloorArea)} m²`, "floor");
      addLine("Špičkový příkon z energetické tabulky",
        [`součet údajů ${formatNumber(estimate.modeledPeakElectricalLoad)} kW u ${formatNumber(estimate.modeledEquipmentCount)} modelových položek včetně základního bydlení`],
        `${formatNumber(estimate.modeledPeakElectricalLoad)} kW`, "volume");
      addLine("Náhradní příkon pro položky bez údajů",
        [`${formatNumber(estimate.fallbackBuildingVolume)} m³ obestavěného prostoru × ${formatNumber(estimate.energyDensity)} W/m³ ÷ 1 000`],
        `${formatNumber(estimate.fallbackPeakElectricalLoad)} kW`, "area");
      addLine("Celkový návrhový příkon",
        [`${formatNumber(estimate.modeledPeakElectricalLoad)} kW z tabulky + ${formatNumber(estimate.fallbackPeakElectricalLoad)} kW náhradní odhad`],
        `${formatNumber(estimate.quantity)} kW`, "volume");
      addLine("Roční elektřina základního bydlení a vybraných vybaveností",
        [`${formatNumber(estimate.modeledAnnualElectricity)} kWh/rok z tabulky + odhad podle poměru špičky ${formatNumber(estimate.peakToAverageRatio)}`],
        `${formatNumber(estimate.annualElectricity)} kWh/rok`, "volume");
      addLine("Délka nového elektrického kabelového vedení", [], `${formatNumber(estimate.networkLength)} m`, "site");
      addLine("Cena kabelového vedení včetně pokládky",
        [`${formatNumber(estimate.networkLength)} m × ${formatCost(estimate.lineRate)}/m × ${formatNumber(estimate.lineSizingRatio)} (poměr celkového a rezidenčního špičkového příkonu)`],
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
      sources.push("Modelový roční odběr elektřiny a špičkový příkon se načítají z energetických sloupců bytového fondu a vybraných vybaveností v listu Model_Vybavenosti. Pro objekty bez kompletních energetických hodnot se používá editovatelný objemový odhad 12 W/m³ (ekvivalent dřívějších 50 W/m² při 4,2 m³/m²) a poměr špičky 4.");
      sources.push("Kapacita a cena trafostanice i kabelové vedení jsou upravitelné odhady, nikoli normové hodnoty nebo RTS ceny.");
      sources.push("Výchozí cena kabelu za metr je odvozena z alokace 30 % scénáře TECH-COST-001 při trase 1 500 m po odečtení pracovního odhadu trafostanice za 3 000 000 Kč; dimenzi, napětí a podmínky připojení musí určit správce a projektant.");
      sources.push("Základní cenu kabelové trasy násobí poměr celkového a rezidenčního špičkového příkonu, aby realizované vybavenosti promítly vyšší návrhové zatížení i do odhadu nákladů. Jde o scénářovou aproximaci, nikoli projektové dimenzování kabelu.");
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
    addLine("Zahrnuté kapacity ordinací",
      [estimate.clinicEstimates.map(({ item: clinic, variant, estimate: clinicEstimate }) =>
        `${displayItemName(clinic)} (${variant === "recommended" ? "doporučená" : "zákonná"} varianta): ${clinicEstimate.metrics[0].value}; ${clinicEstimate.metrics[1].value}`).join(" + ")],
      `${formatNumber(estimate.clinicCount)} ordinací / ambulancí`, "area");
    addLine("Součet investic zahrnutých ordinací",
      [estimate.clinicEstimates.map(({ item: clinic, variant, estimate: clinicEstimate }) =>
        `${displayItemName(clinic)} (${variant === "recommended" ? "doporučená" : "zákonná"}): ${formatNumber(clinicEstimate.cost)} Kč`).join(" + ")],
      formatCost(estimate.grossCost), "rts");
    addLine("Koeficient efektivity společné polikliniky",
      [`${formatCost(estimate.grossCost)} × ${formatNumber(estimate.discountCoefficient)}`],
      formatCost(estimate.cost), "surcharge");
    addLine("Součet užitných ploch před koeficientem",
      [estimate.clinicEstimates.map(({ item: clinic, estimate: clinicEstimate }) =>
        `${displayItemName(clinic)}: ${formatNumber(clinicEstimate.usableArea)} m²`).join(" + ")],
      `${formatNumber(estimate.grossUsableArea)} m²`, "area");
    addLine("Užitná plocha po koeficientu efektivity",
      [`${formatNumber(estimate.grossUsableArea)} m² × ${formatNumber(estimate.discountCoefficient)}`],
      `${formatNumber(estimate.usableArea)} m²`, "area");
    sources.push("Pro každý druh péče se započítá právě jedna varianta: doporučená, pokud existuje; jinak zákonná. Kapacita a počet modelových ordinací se sčítají bez snížení koeficientem.");
    sources.push("Koeficient efektivity je editovatelný uživatelský scénář. Uplatňuje se na souhrnné plochy a investice kvůli sdílení prostor; není to sazba RTS ani metodický normativ.");
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
      const recommendation = estimate.recommendation;
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
      if (recommendation) {
        addLine("Poptávka MMR včetně spádu",
          [`${formatNumber(recommendation.catchmentPopulation)} obyvatel ÷ ${formatNumber(recommendation.populationPerSeat)} obyvatel/místo`],
          `${formatNumber(recommendation.totalDemand)} ${text(getCapacityResultUnit(item), "míst")}`, "area");
        addLine("Ověřená kapacita okolí započtená",
          [recommendation.capacityVerified
            ? `${formatNumber(recommendation.freeCapacity)} volných míst; zdroj doložen; dostupnost ${recommendation.accessRule.summary}`
            : "0 míst: chybí doložený zdroj, volná kapacita nebo vyhovující dostupnost"],
          `${formatNumber(recommendation.appliedCapacity)} ${text(getCapacityResultUnit(item), "míst")}`, "fixed");
        if (recommendation.spatialPointCount !== null) {
          addLine("Orientační minimum pro plošné pokrytí",
            [`${formatNumber(recommendation.spatialPointCount)} bodů × ${formatNumber(recommendation.siteMinimumCapacity)} míst/bod`],
            `${formatNumber(recommendation.spatialCapacityFloor)} ${text(getCapacityResultUnit(item), "míst")}`, "area");
        }
        addLine("Doporučená nová kapacita k zajištění",
          [`max(${formatNumber(recommendation.capacityDemand)} míst kapacitní potřeby, ${formatNumber(recommendation.spatialCapacityFloor)} míst plošného minima)`],
          `${formatNumber(estimate.targetCapacity)} ${text(getCapacityResultUnit(item), "míst")}`, "area");
        sources.push("TB050MMR01, tabulka 1 (tištěná str. 8) a odůvodnění prahů a hustot (tištěné str. 25–26): MŠ 1 dítě/28 obyvatel a práh 420 obyvatel/15 dětí; ZŠ 1. stupně 1 žák/22 obyvatel a práh 2 000 obyvatel/85 žáků (malotřídka 220/10); úplná ZŠ 1 žák 2. stupně/29 obyvatel a práh 4 000 obyvatel/136 žáků.");
        sources.push(`Dostupnost se posuzuje odděleně podle typu území a skutečné vzdálenosti/času: ${recommendation.accessRule.summary}. Plošná hustota je kontrolní ukazatel, ne náhrada skutečné trasy.`);
      } else {
        addLine("Kapacitní poptávka dle modelu",
          [`${formatNumber(estimate.targetCapacity)} míst ÷ ${formatNumber(estimate.modelReferenceDemand)} míst ve zdrojovém modelu`],
          `${formatNumber(estimate.scale, 3)}× rozsahu zdrojového modelu`, "area");
      }
      addLine("Základní užitná plocha", [], `${formatNumber(estimate.baseUsableArea)} m²`, "area");
      addLine("Užitná plocha pro zvolenou kapacitu",
        [`${formatNumber(estimate.baseUsableArea)} m² × ${formatNumber(estimate.targetCapacity)} / ${formatNumber(estimate.modelReferenceDemand)} míst`],
        `${formatNumber(estimate.usableArea)} m²`, "area");
      sources.push("Plocha, objem, areál a investice jsou škálovány poměrem kapacity k velikosti zdrojového modelu; jde o výměrový odhad, nikoli normový projekt budovy.");
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

function renderInfrastructureReserveControls(item) {
  const disclosure = createElement("details", "project-estimate-controls");
  addDisclosureSummary(disclosure, "Upravit výpočet rezervy");
  disclosure.append(createElement("p", "project-estimate-input-note",
    "Rezerva se vypočítá z kapacity a orientační ceny odpovídající základní infrastruktury. Výchozí hodnota 10 % je upravitelný scénářový předpoklad."));
  const fields = createElement("div", "secondary-fields project-estimate-inputs");
  const field = {
    code: "INFRA_RESERVE_PERCENT",
    label: "Rezerva kapacity",
    kind: "percent",
    min: 0,
    max: 1,
    step: 0.01,
    unit: "%",
    color: "surcharge"
  };
  fields.append(createInputControl(field, getInfrastructureReservePercent(item), true,
    `${item.id}-infrastructure`, "itemEstimate", item.key));
  disclosure.append(fields);
  return disclosure;
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
  if (PARKING_SPLIT_ITEM_IDS.has(item.id)) {
    return getProjectPlanningEstimate(item, kind)?.cost ?? null;
  }
  if (item.id === "VYB_04A") {
    return getProjectPlanningEstimate(item, kind)?.cost ?? calc.cost;
  }
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
  const isDelta = kind === "recommended" && !isCompleteStandardMode()
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
    const kind = card.dataset.kind || "legal";
    const projectEstimate = getProjectPlanningEstimate(item, kind);
    if (item.id === "VYB_05B") {
      const hallCount = getItemEstimateQuantity(item);
      quantity.textContent = hallCount > 0
        ? `${formatNumber(hallCount)} ${hallCount === 1 ? "hala" : "haly"}`
        : "bez samostatné haly";
    } else if (projectEstimate?.type === "polyclinic") {
      quantity.textContent = `${formatNumber(projectEstimate.clinicCount)} ordinací`;
    } else {
      const summaryValue = getSummaryRequirementQuantity(calc);
      quantity.textContent = summaryValue === null
        ? "bez počtu" : formatRequirementSummaryAmount(item, summaryValue);
    }
    quantity.textContent = formatRequirementSummaryQuantity(item,
      kind, quantity.textContent);
  }
  const capacitySummary = card.querySelector(".requirement-capacity");
  if (capacitySummary && calc.usesDemandBasedCapacity) {
    capacitySummary.textContent = getCapacitySummaryText(item, calc);
    capacitySummary.hidden = !calc.isCapacityVerified;
  }
}

function hasStandardComparison(item) {
  if (isHealthcareItem(item) || item.id === "VYB_04A") return false;
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
  if (["VYB_09B_1", "VYB_09B_2"].includes(item.id)) return "volných odborných ambulancí";
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
  return isHealthcareDemandItem(item);
}

function isHealthcareDemandItem(item) {
  return HEALTHCARE_DEMAND_ITEM_IDS.has(item.id) && isHealthcareItem(item);
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
  return "Kapacitní poptávka";
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
  const hasUserOverride = hasActiveUserOverride(override);
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
  if (kind === "recommended" && MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) {
    const recommendation = getMmrSchoolRecommendation(item);
    const required = getPlannedQuantity(item, kind);
    if (!recommendation || required === null || required < 0) {
      return {
        required, remaining: null, remainingCapacity: null,
        capacityDemand: recommendation?.totalDemand ?? null,
        capacityUnit: getCapacityUnit(item), usesDemandBasedCapacity: true,
        cost: null, verifiedCapacity: 0, isCapacityVerified: false
      };
    }
    return {
      required,
      remaining: required,
      remainingCapacity: required,
      capacityDemand: recommendation.totalDemand,
      capacityUnit: getCapacityUnit(item),
      usesDemandBasedCapacity: true,
      cost: null,
      verifiedCapacity: recommendation.appliedCapacity,
      availableCapacity: recommendation.capacityVerified ? recommendation.freeCapacity : 0,
      isCapacityVerified: recommendation.capacityVerified || hasUserOverride,
      hasOverride: hasUserOverride
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
    && !hasActiveUserOverride(state.overrides[item.key]);
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
  const costBasis = kind === "recommended" && isMmrWasteYard(item)
    ? item.modelProposed
    : (item.quantityBasis || usesDemandBasedCapacity(item)) ? item.demand : item.modelProposed;
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
    hasOverride: hasActiveUserOverride(state.overrides[item.key])
  };
}

function renderRequirementList(parent, kind, records) {
  parent.replaceChildren();
  const items = getRequirementItems(records, kind);
  const selectedKeys = getRealizationPlan().selectedKeys;
  const visibleItems = state.hideUnrealizedItems
    ? items.filter((item) => selectedKeys.has(item.key))
    : items;
  const specialParkingEstimate = kind === "legal" && state.items.length
    ? getSpecialBuildingParkingEstimate() : null;
  const specialParkingSelected = getRealizationPlan().selectedKeys
    .has(SPECIAL_BUILDING_PARKING_REALIZATION_KEY);
  const showSpecialParking = Boolean(specialParkingEstimate
    && (!state.hideUnrealizedItems || specialParkingSelected));
  const badge = kind === "legal" ? elements.legalCount
    : kind === "recommended" ? elements.recommendedCount
      : kind === "infrastructure" ? elements.infrastructureCount
        : kind === "reserve" ? elements.reserveCount : elements.reviewCount;
  const linkedRules = (state.unlinkedRules[kind] || []).filter((rule) => {
    if (state.hideUnrealizedItems) return false;
    if (state.selectedCategory) return false;
    const query = state.search.toLocaleLowerCase("cs").trim();
    if (!query) return true;
    return [rule[RULE_HEADERS.id], rule[RULE_HEADERS.equipment], rule[RULE_HEADERS.equipmentId],
      rule[RULE_HEADERS.area], rule[RULE_HEADERS.description], rule[RULE_HEADERS.source]]
      .some((value) => String(value ?? "").toLocaleLowerCase("cs").includes(query));
  });
  const itemCount = visibleItems.length + Number(showSpecialParking);
  badge.textContent = linkedRules.length
    ? `Položky: ${itemCount} · ID k přiřazení: ${linkedRules.length}`
    : String(itemCount);
  if (!items.length && !linkedRules.length && !showSpecialParking) {
    const empty = createElement("div", "empty-list", state.hideUnrealizedItems
      ? "V tomto seznamu nejsou položky vybrané k realizaci."
      : kind === "legal"
        ? "V této skupině nebylo nalezeno pravidlo výslovně označené jako právní povinnost."
        : kind === "recommended"
          ? "V této skupině nebylo nalezeno pravidlo výslovně označené jako metodické doporučení."
          : kind === "infrastructure"
            ? "V této skupině nejsou položky technické infrastruktury."
            : kind === "reserve"
              ? "V této skupině nejsou samostatné položky navýšení kapacity."
            : "V této skupině nejsou další nezařazená pravidla.");
    empty.dataset.realizationEmpty = "";
    empty.dataset.requirementKind = kind;
    parent.append(empty);
    return;
  }
  const sortItems = (left, right) => {
    const itemKind = (item) => item.id === RESIDENT_PARKING_RECOMMENDED_ID
      || item.id === RESIDENT_PARKING_RESERVE_ID ? "recommended"
      : isInfrastructureReserveItem(item) ? "infrastructure" : kind;
    const listSatisfied = (item) => itemKind(item) === "recommended"
      && !isCompleteStandardMode()
      && hasStandardComparison(item)
      && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
      ? getStandardComparison(item).delta === 0
      : isRequirementSatisfied(adjustedResult(item, itemKind(item)));
    const leftZero = listSatisfied(left);
    const rightZero = listSatisfied(right);
    return Number(leftZero) - Number(rightZero);
  };
  items.sort(sortItems);
  const groupedItems = new Map();
  for (const item of items) {
    if (!groupedItems.has(item.category)) groupedItems.set(item.category, []);
    groupedItems.get(item.category).push(item);
  }
  const specialParkingCategory = "DOPRAVA, PARKOVÁNÍ A BEZPEČNOST";
  if (kind === "legal" && showSpecialParking
    && !groupedItems.has(specialParkingCategory)) {
    groupedItems.set(specialParkingCategory, []);
  }
  const orderedCategories = [
    ...state.categories.filter((category) => groupedItems.has(category)),
    ...[...groupedItems.keys()].filter((category) => !state.categories.includes(category))
  ];
  for (const category of orderedCategories) {
    const group = createElement("section", "requirement-topic-group");
    const headingRow = createElement("div", "requirement-topic-heading-row");
    const categoryIndex = Math.max(0, state.categories.indexOf(category)) % CATEGORY_COLORS.length;
    const groupToggle = createElement("button",
      `requirement-topic-toggle category-${categoryIndex} shape-${categoryShape(category)}`);
    groupToggle.type = "button";
    groupToggle.dataset.realizationGroupToggle = "";
    groupToggle.setAttribute("aria-label", `Vybrat všechny položky skupiny ${category}`);
    groupToggle.setAttribute("aria-pressed", "false");
    groupToggle.append(createElement("span",
      `category-dot shape-${categoryShape(category)}`, ""));
    headingRow.append(groupToggle);
    headingRow.append(createElement("h4", "requirement-topic-heading", category));
    group.append(headingRow);
    for (const item of groupedItems.get(category)) {
      const itemKind = item.id === RESIDENT_PARKING_RECOMMENDED_ID
        || item.id === RESIDENT_PARKING_RESERVE_ID ? "recommended"
        : isInfrastructureReserveItem(item) ? "infrastructure" : kind;
      group.append(renderItemCard(item, itemKind));
    }
    if (kind === "legal" && category === specialParkingCategory && showSpecialParking) {
      group.append(renderSpecialBuildingParkingCard(specialParkingEstimate));
    }
    const selectableControls = [...group.querySelectorAll("button[data-realization-item-key]")]
    .filter((control) => !control.disabled);
    const groupSelected = selectableControls.length > 0
    && selectableControls.every((control) => control.getAttribute("aria-pressed") === "true");
    groupToggle.setAttribute("aria-pressed", String(groupSelected));
    groupToggle.classList.toggle("is-selected", groupSelected);
    groupToggle.setAttribute("aria-label",
    `${groupSelected ? "Odebrat výběr" : "Vybrat všechny položky"} skupiny ${category}`);
    parent.append(group);
  }
  linkedRules.forEach((rule) => parent.append(renderUnlinkedRuleCard(rule, kind)));
  if (!visibleItems.length && !linkedRules.length && !showSpecialParking) {
    const empty = createElement("div", "empty-list", "V tomto seznamu nejsou položky vybrané k realizaci.");
    empty.dataset.realizationEmpty = "";
    parent.append(empty);
  }
  updateRequirementListVisibility(parent, selectedKeys);
}

function getRequirementItems(records, kind) {
  return records.filter((item) => {
    if (REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(item.id)) return false;
    if (kind === "reserve") return isRequirementReserveItem(item);
    if (isRequirementReserveItem(item)) return false;
    if (item.id === RESIDENT_PARKING_RECOMMENDED_ID) return kind === "review";
    if (item.id === "VYB_01_J") return false;
    if (item.userAdded) return kind === "review"
      || kind === "infrastructure" && isInfrastructureItem(item);
    const hasLegalRule = item.legalRules.some((rule) => !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id])));
    const isInfrastructure = isInfrastructureItem(item);
    const hasPrimaryPlacement = hasLegalRule || item.recommendedRules.length > 0
      || item.id === "VYB_07B" || RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
      || MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id);
    if (kind === "infrastructure") return isInfrastructure;
    if (isInfrastructure) return false;
    if (kind === "legal") {
      if (item.id === "VYB_20") return false;
      return item.id === RESIDENT_PARKING_ID || item.id === "VYB_17B"
        || item.id === "VYB_09D"
        || hasLegalRule && !RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id);
    }
    if (kind === "recommended") return isResidentParkingItem(item)
      ? false
      : item.recommendedRules.length > 0
      && item.id !== "VYB_04A"
      || item.id === "VYB_07B" || RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
      || MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id);
    return !RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
      && item.otherRules.length > 0 && (!hasPrimaryPlacement || REVIEW_ALTERNATIVE_IDS.has(item.id));
  });
}

function isRequirementReserveItem(item) {
  return item.id === RESIDENT_PARKING_RESERVE_ID || isInfrastructureReserveItem(item);
}

function getSpecialBuildingParkingEstimate() {
  const selectedKeys = getRealizationPlan().selectedKeys;
  const selectedItems = state.items
    .filter((item) => selectedKeys.has(item.key))
    .filter((item) => item.id === "VYB_20"
      || SPECIAL_BUILDING_PARKING_RATES.some((rate) => rate.itemIds.includes(item.id)));
  const polyclinicSelected = selectedItems.some((item) => item.id === "VYB_20");
  const nonDuplicatedItems = polyclinicSelected
    ? selectedItems.filter((item) => !["VYB_09A_4", "VYB_09B_1", "VYB_09B_2", "VYB_17", "VYB_18", "VYB_19"].includes(item.id))
    : selectedItems;
  const saved = state.itemEstimateValues[SPECIAL_BUILDING_PARKING_KEY] || {};
  const value = (code, defaultValue) => {
    const stored = numericOrNull(saved[code]);
    return stored === null ? defaultValue : stored;
  };
  let totalFloorArea = 0;
  let totalSalesArea = 0;
  let unestimatedFacilities = 0;
  const facilityCalculations = [];
  const rateResults = SPECIAL_BUILDING_PARKING_RATES.map((rate) => {
    const m2PerStall = value(`RATE_${rate.code}`, rate.m2PerStall);
    const salesAreaShare = value(`SALES_SHARE_${rate.code}`, rate.salesAreaShare ?? 1);
    let area = 0;
    let facilityCount = 0;
    const categoryItems = rate.code === "HEALTH"
      ? [...nonDuplicatedItems.filter((candidate) => rate.itemIds.includes(candidate.id)),
        ...nonDuplicatedItems.filter((candidate) => candidate.id === "VYB_20")]
      : nonDuplicatedItems.filter((candidate) => rate.itemIds.includes(candidate.id));
    for (const item of categoryItems) {
      const kind = item.userAdded ? "review"
        : item.recommendedRules.length || RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
          || MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)
          ? "recommended" : "legal";
      const estimate = getProjectPlanningEstimate(item, kind);
      const itemArea = numericOrNull(estimate?.totalFloorArea)
        ?? numericOrNull(estimate?.grossTotalFloorArea)
        ?? numericOrNull(estimate?.facilityPlan?.totalFloorArea);
      if (itemArea === null) {
        unestimatedFacilities += 1;
        continue;
      }
      const calculationArea = Math.max(0, itemArea) * salesAreaShare;
      const itemStalls = m2PerStall > 0 ? Math.ceil(calculationArea / m2PerStall) : 0;
      area += calculationArea;
      totalSalesArea += rate.code === "RETAIL" ? calculationArea : 0;
      totalFloorArea += Math.max(0, itemArea);
      facilityCount += 1;
      facilityCalculations.push({
        name: displayItemName(item),
        code: rate.code,
        hpp: Math.max(0, itemArea),
        calculationArea,
        salesAreaShare,
        m2PerStall,
        stalls: itemStalls
      });
    }
    return {
      ...rate,
      m2PerStall,
      salesAreaShare,
      area,
      facilityCount,
      stalls: facilityCalculations
        .filter((facility) => facility.code === rate.code)
        .reduce((sum, facility) => sum + facility.stalls, 0)
    };
  });
  const estimatedStalls = rateResults.reduce((sum, rate) => sum + rate.stalls, 0);
  const override = state.overrides[SPECIAL_BUILDING_PARKING_REALIZATION_KEY];
  const overrideQuantity = numericOrNull(override?.quantity);
  const stalls = hasActiveUserOverride(override) ? overrideQuantity : estimatedStalls;
  const parkingSource = state.items.find((item) => item.id === RESIDENT_PARKING_ID);
  const parkingValue = (code, fallback) => numericOrNull(saved[code]) ?? (parkingSource
    ? numericOrNull(getParkingEstimateValue(parkingSource, "legal", code)) : null) ?? fallback;
  const buildingShare = Math.min(1, Math.max(0, value("BUILDING_SHARE", parkingValue("PARKING_BUILDING_SHARE", 0.2))));
  const volumePerStall = Math.max(0, value("VOLUME_PER_STALL_M3", parkingValue("PARKING_VOLUME_PER_STALL_M3", 18)));
  const buildingRate = Math.max(0, value("BUILDING_RATE_M3", parkingValue("PARKING_BUILDING_RATE_M3", 8995)));
  const surfaceRate = Math.max(0, value("SURFACE_RATE_M3", parkingValue("PARKING_SURFACE_RATE_M3", 7500)));
  const surfaceAreaPerStall = Math.max(0, value("SURFACE_AREA_PER_STALL_M2", parkingValue("PARKING_SURFACE_AREA_PER_STALL_M2", 25)));
  const surcharge = Math.max(0, value("COST_SURCHARGE", parkingValue("ITEM_COST_SURCHARGE", 0.08)));
  const buildingStalls = stalls * buildingShare;
  const surfaceStalls = stalls - buildingStalls;
  const pavedArea = surfaceStalls * surfaceAreaPerStall;
  const zrn = Math.round(buildingStalls * volumePerStall * buildingRate
    + surfaceStalls * volumePerStall * surfaceRate);
  return {
    selectedItems: nonDuplicatedItems,
    facilityCount: rateResults.reduce((sum, rate) => sum + rate.facilityCount, 0),
    unestimatedFacilities,
    totalFloorArea,
    totalSalesArea,
    rateResults,
    facilityCalculations,
    estimatedStalls,
    hasUserOverride: hasActiveUserOverride(override),
    overrideQuantity,
    stalls,
    buildingShare,
    buildingStalls,
    surfaceStalls,
    volumePerStall,
    buildingRate,
    surfaceRate,
    surfaceAreaPerStall,
    pavedArea,
    surcharge,
    zrn,
    cost: Math.round(zrn * (1 + surcharge))
  };
}

function renderSpecialBuildingParkingCard(estimate) {
  const card = createElement("details", "requirement-card special-parking-card");
  card.dataset.kind = "legal";
  card.dataset.itemKey = SPECIAL_BUILDING_PARKING_REALIZATION_KEY;
  card.dataset.specialParkingFacilityCount = String(estimate.facilityCount);
  const selected = getRealizationPlan().selectedKeys
    .has(SPECIAL_BUILDING_PARKING_REALIZATION_KEY);
  const nearbyCapacity = numericOrNull(
    state.capacityEdits[SPECIAL_BUILDING_PARKING_REALIZATION_KEY]?.quantity);
  const capacitySatisfied = state.capacityEdits[
    SPECIAL_BUILDING_PARKING_REALIZATION_KEY]?.satisfiesNearbyCapacity === true
    || nearbyCapacity !== null && nearbyCapacity >= estimate.stalls;
  const isZero = estimate.stalls === 0;
  card.classList.toggle("is-zero", isZero);
  card.classList.toggle("is-capacity-satisfied", capacitySatisfied);
  card.classList.toggle("is-realization-selected", selected);
  card.open = false;
  const summary = document.createElement("summary");
  summary.className = "special-parking-summary";
  const top = createElement("div", "requirement-card-top");
  const realizationButton = createElement("button",
    `category-dot-control${selected ? " is-selected" : ""}`);
  realizationButton.type = "button";
  realizationButton.dataset.realizationItemKey = SPECIAL_BUILDING_PARKING_REALIZATION_KEY;
  realizationButton.disabled = capacitySatisfied;
  realizationButton.setAttribute("aria-pressed", String(selected));
  realizationButton.setAttribute("aria-label",
    `${selected ? "Zrušit výběr" : "Vybrat"} k realizaci: Parkování realizovaných specializovaných staveb`);
  realizationButton.title = selected ? "Odebrat z realizace" : "Vybrat k realizaci";
  if (capacitySatisfied) realizationButton.title = "Kapacita okolí je označena jako dostatečná.";
  realizationButton.append(createElement("span", "category-dot shape-circle"));
  top.append(realizationButton);
  const title = createElement("span", "requirement-title");
  title.append(createElement("strong", "special-parking-title",
    "Parkování realizovaných specializovaných staveb (splnění vyhlášky č. 146/2024 Sb.)"));
  top.append(title);
  const summaryMetrics = createElement("span", "requirement-summary-metrics");
  summaryMetrics.append(createElement("span", "requirement-quantity",
    formatValue(estimate.stalls, "stání")));
  summaryMetrics.append(createElement("small", "requirement-investment", formatCost(estimate.cost)));
  top.append(summaryMetrics, createElement("span", "expand-mark", "＋"));
  summary.append(top);
  card.append(summary);
  const content = createElement("div", "special-parking-content");
  content.append(createElement("p", "special-parking-note",
    "Odhad sčítá pouze nové parkování pro vybavenosti označené k realizaci; nezahrnuje rezidentské parkování. Počty se odvozují z HPP jednotlivých realizovaných objektů a editovatelných koeficientů. Pro obchody je výchozí pravidlo 1 stání / 30 m² prodejní HPP a podíl prodejní plochy z HPP je nastaven na 80 %. Koeficienty ostatních typů jsou scénářové předpoklady, nikoli sazby stanovené vyhláškou; ověřte konkrétní stavbu podle přílohy č. 1 vyhlášky č. 146/2024 Sb."));
  const results = createElement("section", "special-parking-results");
  results.append(createElement("h4", "", RESULTS_HEADING));
  const metrics = createElement("div", "special-parking-metrics");
  const addMetric = (label, value, unit = "") => {
    const metric = createElement("div", "item-result-metric");
    metric.append(createElement("span", "", label));
    metric.append(createElement("strong", "", formatValue(value, unit)));
    metrics.append(metric);
  };
  addMetric("Realizované budovy s použitelnou výměrou", estimate.facilityCount, "budov");
  addMetric("Celková HPP realizovaných budov", estimate.totalFloorArea, "m²");
  addMetric("Započtená HPP prodejní plochy obchodů", estimate.totalSalesArea, "m²");
  addMetric("Modelová potřeba parkovacích stání", estimate.estimatedStalls, "stání");
  if (estimate.hasUserOverride) {
    addMetric("Uživatelský návrh parkovacích stání", estimate.stalls, "stání");
  }
  addMetric("Z toho parkovací dům / venkovní", `${formatNumber(estimate.buildingStalls)} / ${formatNumber(estimate.surfaceStalls)}`, "stání");
  addMetric("Orientační zpevněná plocha", estimate.pavedArea, "m²");
  addMetric("Orientační investice včetně přirážky", estimate.cost, "Kč");
  results.append(metrics);
  content.append(results);
  content.append(renderSpecialParkingCapacityEditor(estimate));
  const rateDetails = createElement("details", "special-parking-inputs");
  addDisclosureSummary(rateDetails, "Koeficienty a cena – editovatelné předpoklady");
  const rateFields = createElement("div", "special-parking-rate-fields");
  estimate.rateResults.forEach((rate) => {
    const field = createElement("label", "special-parking-rate");
    const label = createElement("span", "", `${rate.label} – m² HPP na stání`);
    const input = createElement("input", "");
    input.type = "number";
    input.min = "1";
    input.step = "1";
    input.value = String(numericOrNull(state.itemEstimateValues[SPECIAL_BUILDING_PARKING_KEY]?.[`RATE_${rate.code}`]) ?? rate.m2PerStall);
    input.dataset.specialParkingInput = `RATE_${rate.code}`;
    field.append(label, input);
    rateFields.append(field);
    if (rate.code === "RETAIL") {
      const shareField = createElement("label", "special-parking-rate");
      shareField.append(createElement("span", "", "Podíl prodejní plochy z HPP obchodů"));
      const shareInput = createElement("input", "");
      shareInput.type = "number";
      shareInput.min = "0";
      shareInput.max = "100";
      shareInput.step = "1";
      shareInput.value = String((numericOrNull(state.itemEstimateValues[SPECIAL_BUILDING_PARKING_KEY]?.SALES_SHARE_RETAIL)
        ?? rate.salesAreaShare) * 100);
      shareInput.dataset.specialParkingInput = "SALES_SHARE_RETAIL";
      shareInput.dataset.specialParkingScale = "100";
      shareField.append(shareInput, createElement("small", "", "%"));
      rateFields.append(shareField);
    }
  });
  const costInputs = [
    { code: "BUILDING_SHARE", label: "Podíl parkovací dům", value: estimate.buildingShare * 100, min: 0, max: 100, step: 1, unit: "%" },
    { code: "VOLUME_PER_STALL_M3", label: "Objem domu na stání", value: estimate.volumePerStall, min: 0, step: 1, unit: "m³/stání" },
    { code: "BUILDING_RATE_M3", label: "Sazba parkovacího domu", value: estimate.buildingRate, min: 0, step: 100, unit: "Kč/m³ OP" },
    { code: "SURFACE_RATE_M3", label: "Sazba venkovního stání", value: estimate.surfaceRate, min: 0, step: 100, unit: "Kč/m³ ekv." },
    { code: "SURFACE_AREA_PER_STALL_M2", label: "Zpevněná plocha na venkovní stání", value: estimate.surfaceAreaPerStall, min: 0, step: 1, unit: "m²/stání" },
    { code: "COST_SURCHARGE", label: "VRN a rozpočtová přirážka", value: estimate.surcharge * 100, min: 0, step: 1, unit: "%" }
  ];
  costInputs.forEach((fieldData) => {
    const field = createElement("label", "special-parking-rate");
    const label = createElement("span", "", fieldData.label);
    const input = createElement("input", "");
    input.type = "number";
    input.min = String(fieldData.min);
    if (fieldData.max !== undefined) input.max = String(fieldData.max);
    input.step = String(fieldData.step);
    input.value = String(fieldData.value);
    input.dataset.specialParkingInput = fieldData.code;
    input.dataset.specialParkingScale = fieldData.unit === "%" ? "100" : "1";
    field.append(label, input, createElement("small", "", fieldData.unit));
    rateFields.append(field);
  });
  rateDetails.append(rateFields);
  content.append(rateDetails);
  const calculation = createElement("section", "special-parking-calculation");
  calculation.append(createElement("h4", "", "Postup výpočtu"));
  calculation.append(createElement("p", "special-parking-formula",
    "1) U každé realizované budovy se vezme vypočtená HPP. 2) U obchodu se HPP vynásobí podílem prodejní plochy. 3) Parkovací stání pro každý objekt = zaokrouhlení nahoru (započtená plocha ÷ koeficient m²/stání). 4) Potřeba sečte stání jednotlivých objektů."));
  const facilityList = createElement("ul", "special-parking-facility-list");
  estimate.facilityCalculations.forEach((facility) => {
    const areaExpression = facility.code === "RETAIL"
      ? `${formatNumber(facility.hpp)} m² HPP × ${formatNumber(facility.salesAreaShare * 100)} % = ${formatNumber(facility.calculationArea)} m² prodejní plochy`
      : `${formatNumber(facility.hpp)} m² HPP`;
    facilityList.append(createElement("li", "",
      `${facility.name}: ${areaExpression}; ${formatNumber(facility.calculationArea)} ÷ ${formatNumber(facility.m2PerStall)} m²/stání → ${facility.stalls} stání`));
  });
  if (!estimate.facilityCalculations.length) {
    facilityList.append(createElement("li", "", "Zatím není vybraná žádná realizovaná budova s vypočtenou HPP."));
  }
  calculation.append(facilityList);
  calculation.append(createElement("p", "special-parking-formula",
    `Ocenění: ${formatNumber(estimate.buildingStalls)} stání v parkovacím domě × ${formatNumber(estimate.volumePerStall)} m³/stání × ${formatCost(estimate.buildingRate)}/m³ + ${formatNumber(estimate.surfaceStalls)} venkovních stání × ${formatNumber(estimate.volumePerStall)} m³/stání × ${formatCost(estimate.surfaceRate)}/m³; poté přirážka ${formatNumber(estimate.surcharge * 100)} % = ${formatCost(estimate.cost)}.`));
  content.append(calculation);
  if (estimate.unestimatedFacilities) {
    content.append(createElement("p", "special-parking-warning",
      `${estimate.unestimatedFacilities} realizovaných položek nemá vypočtenou HPP, a proto zatím nejsou do odhadu parkování zahrnuty.`));
  }
  content.append(createFeedbackButton(
    "Parkování realizovaných specializovaných staveb (splnění vyhlášky č. 146/2024 Sb.)"
  ));
  card.append(content);
  return card;
}

function renderSpecialParkingCapacityEditor(estimate) {
  const disclosure = createElement("details", "item-section capacity-section");
  addDisclosureSummary(disclosure, "Ověřený stav v okolí a uživatelský návrh");
  disclosure.append(renderAvailabilitySearch({
    id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY,
    name: "Parkování realizovaných specializovaných staveb"
  }));
  const editor = createElement("section", "capacity-editor");
  const heading = createElement("div", "capacity-heading");
  heading.append(createElement("span", "capacity-unit", "Jednotka pro ověření: stání"));
  editor.append(heading);
  const controls = createElement("div", "capacity-controls");
  const capacity = state.capacityEdits[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] || {};
  const override = state.overrides[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] || {};
  const addNumberField = (label, value, field, placeholder) => {
    const wrapper = createTextInput(label, value, {
      id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY,
      field,
      kind: "number"
    }, placeholder);
    const input = wrapper.querySelector("input");
    input.type = "number";
    input.min = "0";
    input.step = "any";
    controls.append(wrapper);
  };
  addNumberField("Celková ověřená kapacita v okolí (stání)",
    capacity.quantity, "quantity", "Prázdné = neověřeno");
  controls.append(createTextInput("Zdroj / ověření (nepovinné)",
    capacity.source, { id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY, field: "source" },
    "Např. GIS, plán nebo správce"));
  controls.append(createTextInput("Datum (doplní se automaticky)",
    capacity.date, { id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY, field: "date" }));
  controls.append(createTextInput("Poznámka k ověření",
    capacity.note, { id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY, field: "note" },
    "Rozsah, kapacita a omezení"));
  addNumberField("Uživatelsky upravený návrh (stání)",
    override.quantity, "override", "Ponechat modelovou hodnotu");
  controls.append(createTextInput("Autor přepisu (nepovinné)",
    override.author, { id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY, field: "overrideAuthor" },
    "Volitelné"));
  controls.append(createTextInput("Důvod / poznámka k úpravě návrhu (nepovinné)",
    override.note, { id: SPECIAL_BUILDING_PARKING_REALIZATION_KEY, field: "overrideNote" },
    "Volitelné"));
  editor.append(controls);
  const verifiedCapacity = numericOrNull(capacity.quantity);
  const remaining = verifiedCapacity === null ? null
    : Math.max(0, estimate.stalls - verifiedCapacity);
  editor.append(createElement("p", "capacity-result",
    verifiedCapacity === null
      ? `Modelová potřeba činí ${formatValue(estimate.stalls, "stání")}; ověřená kapacita v okolí zatím není zadána.`
      : `Modelová potřeba činí ${formatValue(estimate.stalls, "stání")}, ověřená kapacita v okolí ${formatValue(verifiedCapacity, "stání")}; zbývající potřeba je ${formatValue(remaining, "stání")}.`));
  editor.append(createElement("p", "capacity-caution",
    "Zadání kapacity, zdroje, poznámky, autora i uživatelského návrhu je dobrovolné. Výpočet parkování se řídí zvolenými koeficienty; uživatelský návrh přepočítá množství i cenu."));
  editor.append(createElement("p", "capacity-caution capacity-audit-summary",
    getUserOverrideAuditSummary(override, "stání")));
  disclosure.append(editor);
  return disclosure;
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
    card.hidden = state.hideUnrealizedItems && (card.classList.contains("special-parking-card")
      ? !selectedKeys.has(SPECIAL_BUILDING_PARKING_REALIZATION_KEY)
      : !selectedKeys.has(card.dataset.itemKey));
  }
  list.querySelectorAll(".reserve-capacity-group").forEach((group) => {
    group.hidden = !group.querySelector(".requirement-card:not([hidden])");
  });
  const visibleCards = cards.filter((card) => !card.hidden).length;
  const hasOtherContent = Boolean(list.querySelector(".unlinked-rule-card"));
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
  for (const [list, badge] of [
    [elements.legalList, elements.legalCount],
    [elements.infrastructureList, elements.infrastructureCount],
    [elements.recommendedList, elements.recommendedCount],
    [elements.reviewList, elements.reviewCount],
    [elements.reserveList, elements.reserveCount]
  ]) {
    const visibleCards = updateRequirementListVisibility(list, selectedKeys);
    const linkedRules = list.querySelectorAll(".unlinked-rule-card").length;
    badge.textContent = linkedRules
      ? `Položky: ${visibleCards} · ID k přiřazení: ${linkedRules}`
      : String(visibleCards);
  }
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
  body.append(createFeedbackButton(
    `Nepřiřazené pravidlo: ${displayName} [${rule[RULE_HEADERS.id]}] · seznam: ${kind}`
  ));
  details.append(body);
  return details;
}

function displayRuleName(rule) {
  const id = String(rule[RULE_HEADERS.id] ?? "").trim();
  const names = {
    "ZDR-LEG-001": "Časová dostupnost hrazené zdravotní péče – NV č. 307/2012 Sb.",
    "ZDR-SPEC-001": "Časová dostupnost specializované zdravotní péče – NV č. 307/2012 Sb.",
    "ZDR-MET-001": "Dostupnost ambulantní zdravotní péče – metodický standard",
    "ZDR-MMR-GROUP-2": "Ambulantní specialisté – skupina 2 MMR",
    "ZDR-MMR-GROUP-3": "Ambulantní specialisté – skupina 3 MMR",
    "ZEL-MMR-PARK-001": "Veřejná parková zeleň – plošný a docházkový standard MMR",
    "ZEL-LEG-001": "Veřejné prostranství – vyhláška č. 501/2006 Sb., § 7 odst. 2",
    "ZEL-MET-001": "Předškolní hřiště – docházková dostupnost",
    "ZEL-MET-002": "Hřiště pro mladší děti – docházková dostupnost",
    "ZEL-MET-003": "Hřiště pro mládež a dospělé – skutečná pěší docházka",
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
  return [RESIDENT_PARKING_RECOMMENDED_ID, RESIDENT_PARKING_RESERVE_ID].map((id) => ({
    ...source,
    id,
    key: `${source.key}::${id}`,
    parkingSourceKey: source.key,
    name: id === RESIDENT_PARKING_RECOMMENDED_ID
      ? "Rezidentské parkování (dřívější režim – vyhláška č. 268/2009 Sb.)"
      : "Rezidentské parkování – navýšení kapacity",
    legalRules: [],
    recommendedRules: [],
    otherRules: [],
    quantityBasis: false,
    proposed: 0,
    modelProposed: 0,
    demand: 0
  }));
}

function createInfrastructureReserveVariants(items) {
  return items.filter((item) => INFRASTRUCTURE_ITEM_IDS.has(item.id)).map((source) => ({
    ...source,
    id: `${source.id}${INFRASTRUCTURE_RESERVE_SUFFIX}`,
    key: `${source.key}::${source.id}${INFRASTRUCTURE_RESERVE_SUFFIX}`,
    name: `${displayItemName(source)} – rezerva kapacity`,
    infrastructureReserveSourceKey: source.key,
    userAdded: false,
    legalRules: [],
    recommendedRules: [],
    otherRules: [],
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
  const summary = [];
  if (travel !== null && limit !== null) {
    summary.push(travel <= limit
      ? `Zadaný dojezd ${formatNumber(travel)} min nepřekračuje vložený limit ${formatNumber(limit)} min.`
      : `Zadaný dojezd ${formatNumber(travel)} min překračuje vložený limit ${formatNumber(limit)} min.`);
  }
  if (capacity !== null && demand !== null) {
    const difference = capacity - demand;
    summary.push(difference >= 0
      ? `Zadaná kapacita převyšuje zadanou potřebu o ${formatNumber(difference)} osob.`
      : `Zadaný deficit je ${formatNumber(Math.abs(difference))} osob.`);
  } else {
    summary.push("Kapacitní rezerva/deficit nelze vypočítat bez potřeby i kapacity.");
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
    isHealthcareDemandItem(item)
      ? "Výše uvedený populační model navrhuje potřebnou souč. zátěž a počet ambulancí. Zde ověřte jednotlivé odbornosti, poskytovatele, skutečnou dostupnost a volnou kapacitu; modelové poměry jsou upravitelné plánovací předpoklady, nikoli právní norma."
      : "Záznam neznamená automatickou povinnost vybudovat ordinaci. Ověřte poskytovatele, dostupnost a kapacitu v registrech, u pojišťoven a přímo u pracoviště. Limity a skupiny ověřte v citovaném zdroji pro konkrétní službu."));
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
  if (item.id === "VYB_17A") return "Služebna Policie ČR";
  if (TECHNICAL_CATALOG_NAMES[item.id]) return TECHNICAL_CATALOG_NAMES[item.id];
  if (REVIEW_ALTERNATIVE_LABELS[item.id]) return REVIEW_ALTERNATIVE_LABELS[item.id].name;
  if (item.id === "VYB_20") return "Poliklinika / Sdružené Ordinace";
  if (item.id === "VYB_16A") return "Hasičská stanice HZS – pokrytí a případná potřeba";
  if (item.id === "VYB_16B") return "Jednotka požární ochrany – ruční ověření pokrytí";
  if (item.id === RESIDENT_PARKING_ID) return "Rezidentské parkování (splnění vyhlášky č. 146/2024 Sb.)";
  if (item.id === RESIDENT_PARKING_LAW_ID) return "Rezidentské parkování – zákonná varianta";
  if (item.id === RESIDENT_PARKING_RECOMMENDED_ID) {
    return "Rezidentské parkování (dřívější režim – vyhláška č. 268/2009 Sb.)";
  }
  if (item.id === RESIDENT_PARKING_RESERVE_ID) return "Rezidentské parkování – navýšení kapacity";
  if (isInfrastructureReserveItem(item)) {
    const source = getInfrastructureReserveSource(item);
    const sourceName = source ? displayItemName(source) : cleanDisplayName(item.name);
    const conciseSourceName = sourceName
      .replace(/\s*[–-]\s*cena za m³ objemu(?:\s*[–-]\s*)?/i, " ")
      .replace(/\s+/g, " ")
      .trim();
    return `${conciseSourceName} – navýšení kapacity`;
  }
  let name = cleanDisplayName(ITEM_DISPLAY_NAMES[item.id] || item.name)
    .replace(/\s*\(NV č\. 307\/2012 Sb\.\)/gi, "")
    .trim();
  if (["VYB_01", "VYB_02A", "VYB_02B", "VYB_02C_1", "VYB_02C_2"].includes(item.id)) {
    name = name.replace(/\s*\((?:MŠ|ZŠ(?:\s*[12])?|SŠ\s*\/\s*Gym|ZUŠ)\)\s*$/i, "").trim();
  }
  const assumption = PARKING_SHARE_ASSUMPTIONS[item.id];
  return assumption
    ? `${name} (odhad pro ${formatNumber(getParkingShare(item))} % obyvatel)`
    : name;
}

function getItemBriefRule(item, kind, rules) {
  if (isInfrastructureReserveItem(item)) {
    const source = getInfrastructureReserveSource(item);
    return `Samostatné navýšení o ${formatNumber(getInfrastructureReservePercent(item) * 100)} % z vypočtené kapacity ${source ? displayItemName(source) : "základní infrastruktury"}; při výběru se základní položka této infrastruktury nerealizuje.`;
  }
  if (item.id === RESIDENT_PARKING_RESERVE_ID) {
    return "Samostatné navýšení se počítá pouze z jedné zvolené základní varianty parkování; základní varianty se vzájemně vylučují.";
  }
  if (item.id === RESIDENT_PARKING_LAW_ID) {
    return "Úplná zákonná alternativa podle vyhlášky č. 146/2024 Sb.; její počet je stejný jako v zákonné položce. Zvolte ji místo doporučené varianty, nesčítá se s ní.";
  }
  if (item.id === RESIDENT_PARKING_RECOMMENDED_ID) {
    return "Úplná samostatná alternativa k zákonnému minimu; není to rozdíl ani přírůstek k zákonné variantě. Jde o transparentní historický model podle starší úpravy, nikoli o platné zákonné minimum. Vyberte tuto variantu místo zákonné; její kapacita se nesčítá.";
  }
  if (item.id === RESIDENT_PARKING_ID) {
    return "Zákonné minimum podle vyhlášky č. 146/2024 Sb.; samostatná zákonná položka, v doporučeném seznamu se neduplikuje.";
  }
  if (item.id === "VYB_05B") {
    return "Uživatelská scénářová položka: MMR ani zákon pro ni nestanoví požadovaný počet ani povinnost stavby; výchozí volba pouze modeluje samostatnou novou halu.";
  }
  if (item.id === "VYB_17A") {
    return "Doporučení interních metodik policie (nikoli MMR ani zákonný standard): typ A/B – zásah hlídky do 10 minut; typ C – do 15 minut; typ D – požadavek se neměří. Modelový počet policistů a služeben je pouze scénářový a musí se ověřit podle stávajícího pokrytí.";
  }
  if (item.id === "VYB_18A") {
    return "Počet stanovišť je kruhový síťový odhad, nikoli zákonná kvóta.";
  }
  if (item.id === "VYB_18B") {
    return "Vzdálenost 150 m je scénářový předpoklad; počet hnízd ověřte s obcí a svozovou společností.";
  }
  if (item.id === "VYB_07B") {
    return "Uživatelský scénář městského odkupu bytů; nejde o zákonnou ani metodickou povinnost. Výpočet vychází z bytové skladby projektu a zadaného podílu odkupu.";
  }
  if (item.id === "VYB_09D") {
    return "Zákonná dostupnost ZZS: výjezdová skupina má dosáhnout místa události nejpozději do 20 minut od převzetí pokynu k výjezdu, s výjimkami podle § 5 zákona č. 374/2011 Sb. Kontrola dojezdu sama neurčuje povinnost postavit novou základnu; ověřte krajský plán ZZS a stávající kapacitu.";
  }
  if (item.id === "VYB_16A") {
    return "Ověřte skutečný zásahový dojezd hasičské stanice: v typech A/B orientačně 5–7 minut (modelová horní mez 7 minut), v typech C/D 10–15 minut (horní mez 15 minut). Rozhoduje požární poplachový plán a HZS; jde o silniční zásahový dojezd, nikoli pěší dostupnost.";
  }
  if (item.id === VOLUNTEER_FIRE_STATION_ID) {
    return "Doporučený standard TAČR/MMR: skutečná pěší trasa k hasičské zbrojnici dobrovolných hasičů do 1 000 m (tabulka 1, tištěná str. 12). Pěší dostupnost neprokazuje kategorii JPO, akceschopnost ani čas zásahu; výměry budovy a cena jsou modelovány na polovinu výchozího scénáře HZS.";
  }
  if (item.id === HEATING_BOILER_ID) {
    return "Scénářová položka vložená uživatelem bez samostatného zákonného či metodického prahu; výkon a cena kotelny jsou editovatelné odhady, které je nutné ověřit energetickým posudkem a projektem.";
  }
  if (kind === "recommended" && item.socialAvailabilityStandard) {
    return `Doporučený standard: ${getSocialAvailabilityStandardDescription(item.socialAvailabilityStandard)} (metodika MMR ČR, tabulka 1, tištěná str. 9).`;
  }
  if (kind === "recommended" && ["VYB_03A", "VYB_03B", "VYB_03C"].includes(item.id)) {
    const standardText = {
      VYB_03A: "Typ A/B/C/D – skutečná pěší docházka do 200 m.",
      VYB_03B: "Typ A/B/C/D – skutečná pěší docházka do 500 m.",
      VYB_03C: "Typ A/B/C – skutečná pěší docházka do 500 m; typ D – skutečná pěší docházka do 1 000 m."
    }[item.id];
    return `Doporučený standard: ${standardText} Ověřuje se skutečná pěší trasa, nikoli vzdušná vzdálenost (metodika MMR ČR, tabulka Hřiště, tištěná str. 12).`;
  }
  if (kind === "recommended" && item.id === "VYB_02C_1") {
    return "Doporučený standard: Typ A/B – dostupnost veřejnou dopravou do 45 minut; sídelní přítomnost je doporučena v typech A/B při více než 20 000 obyvatelích. Prah se posuzuje ze současných obyvatel řešeného území a přírůstku záměru; kapacita nové školy se odhaduje z přírůstku. (Metodika MMR ČR, tabulka 1, tištěná str. 8.)";
  }
  if (kind === "recommended" && item.id === "VYB_02C_2") {
    return "Doporučený standard: Typ A – sídelně-strukturální přítomnost v obci; typ B/C, obce > 5 000 obyvatel – doporučená sídelně-strukturální přítomnost v obci. Prah se posuzuje ze současných obyvatel řešeného území a přírůstku záměru; kapacita nové ZUŠ se odhaduje z přírůstku. (Metodika MMR ČR, tabulka 1, tištěná str. 8.)";
  }
  if (kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
    const standard = MMR_SCHOOL_STANDARDS[item.id];
    const access = getMmrSchoolAccessRule(item);
    const threshold = standard.smallSchoolPopulation
      ? `běžný práh ${formatNumber(standard.minimumPopulation)} obyvatel/${formatNumber(standard.minimumCapacity)} míst; malotřídka od ${formatNumber(standard.smallSchoolPopulation)} obyvatel/${formatNumber(standard.smallSchoolCapacity)} míst`
      : `orientační minimum ${formatNumber(standard.minimumPopulation)} obyvatel/${formatNumber(standard.minimumCapacity)} míst`;
    return `Doporučení MMR: ${formatNumber(standard.populationPerSeat)} obyvatel na místo; ${threshold}. Dostupnost: ${access.summary}.`;
  }
  if (isHealthcareItem(item)) {
    if (kind === "legal" && item.legalRules.some((rule) =>
      ["ZDR-LEG-001", "ZDR-SPEC-001"].includes(String(rule[RULE_HEADERS.id])))) {
      const specialty = HEALTH_SPECIALTIES[item.id] || [];
      const limits = [...new Set(specialty.map((entry) => entry.limitMinutes).filter(Number.isFinite))];
      const limitText = limits.length === 1 ? `${limits[0]} min` : limits.length > 1
        ? `${Math.min(...limits)}–${Math.max(...limits)} min podle odbornosti`
        : "podle konkrétní hrazené služby";
      return `Právní dostupnost hrazené péče: dojezd nejvýše ${limitText}; nejde o povinnost zřídit ordinaci (NV č. 307/2012 Sb., příloha č. 1).`;
    }
    if (kind === "recommended" && item.recommendedRules.length) {
      if (["VYB_09A_4", "VYB_16", "VYB_17", "VYB_18", "VYB_19", "VYB_20"].includes(item.id)) {
        return "Doporučený standard: Typ A/B/C, obce > 2 000 obyvatel – sídelně-strukturální přítomnost v obci a skutečná pěší dostupnost 600 m; typ D – dojezdová doba do 35 minut. (Metodika MMR ČR, tabulka 1, tištěná str. 9; dojezd dle NV č. 307/2012 Sb.)";
      }
      if (["VYB_09B_1", "VYB_09B_2"].includes(item.id)) {
        const group = item.id === "VYB_09B_1" ? "2" : "3";
        const neurologyNote = item.id === "VYB_09B_2"
          ? " Neurologii zařaďte až po ověření konkrétní služby."
          : "";
        return `Doporučený standard: Typ A – sídelně-strukturální přítomnost v obci; typ B, obce > 5 000 obyvatel – doporučená přítomnost v obci (metodika MMR ČR, tabulka 1, skupina ${group}, tištěná str. 10).${neurologyNote}`;
      }
    }
  }

  if (kind === "recommended" && hasStandardComparison(item)
    && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
    && !isCompleteStandardMode()) {
    return "Doporučený scénář zobrazuje celou zvolenou variantu, nikoli pouze přírůstek nad zákonné minimum.";
  }
  if (kind === "recommended" && isCompleteStandardMode()
    && hasStandardComparison(item) && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
    return "Ve standardním režimu je požadavek porovnán s právním minimem; započítá se pouze výsledná potřeba zvoleného standardu.";
  }

  const applicableRules = kind === "legal" ? item.legalRules
    : kind === "recommended" ? item.recommendedRules : item.otherRules;
  const rulePool = [...rules, ...applicableRules]
    .filter((entry, index, entries) => entries.findIndex((candidate) =>
      reviewKey(candidate) === reviewKey(entry)) === index);
  const rule = rulePool.find((entry) => entry[RULE_HEADERS.condition] || entry[RULE_HEADERS.description])
    || rulePool[0];
  const detail = rule
    ? String(rule[RULE_HEADERS.condition] || rule[RULE_HEADERS.description]
      || rule[RULE_HEADERS.output] || "").replace(/\s+/g, " ").trim()
    : "";
  if (detail) {
    const concise = detail.length > 250 ? `${detail.slice(0, 247).trimEnd()}…` : detail;
    const prefix = kind === "legal" ? "Zákonná podmínka:" : kind === "recommended"
      ? "Doporučený standard:" : "Kontrolní pravidlo:";
    const rawSource = String(rule?.[RULE_HEADERS.source] ?? "").replace(/\s+/g, " ").trim();
    const sourceReference = /TB050MMR0*1|TAČR\/MMR|standardy dostupnosti veřejné infrastruktury/i.test(rawSource)
      ? "metodika MMR ČR"
      : /307\/2012/i.test(rawSource) ? "NV č. 307/2012 Sb."
        : /561\/2004/i.test(rawSource) ? "školský zákon č. 561/2004 Sb."
          : rawSource;
    const location = String(rule?.[RULE_HEADERS.location] ?? "").replace(/\s+/g, " ").trim();
    const sourceParts = [sourceReference, location && sourceReference !== rawSource ? location : ""]
      .filter(Boolean);
    const sourceCitation = sourceParts.join(", ");
    const shortSource = sourceCitation.length > 120
      ? `${sourceCitation.slice(0, 117).trimEnd()}…` : sourceCitation;
    return `${prefix} ${concise}${shortSource ? ` (${shortSource})` : ""}`;
  }
  if (kind === "legal") return "Pro tuto položku není přiřazena samostatná zákonná kapacitní kvóta.";
  if (kind === "recommended") return "Pro tuto položku není přiřazen samostatný metodický práh; ověřte místní potřebu a kapacitu.";
  return "Projektová položka bez samostatného zákonného nebo metodického prahu.";
}

function displayRequirementItemName(item, kind) {
  const name = displayItemName(item);
  if (item.id === "VYB_05B") return name;
  if (item.id === "VYB_04A") {
    return kind === "recommended"
      ? `${name} (metodika MMR ČR)`
      : `${name} (vyhláška č. 501/2006 Sb., § 7 odst. 2)`;
  }
  if (item.id === "VYB_04B" && kind === "recommended") {
    return `${name} (metodika MMR ČR)`;
  }
  if (isHealthcareItem(item) && item.legalRules.some((rule) =>
    ["ZDR-LEG-001", "ZDR-SPEC-001"].includes(String(rule[RULE_HEADERS.id]))) && kind === "legal") {
    return `${name} (NV č. 307/2012 Sb., příloha č. 1)`;
  }
  if (isHealthcareItem(item) && kind === "recommended" && item.recommendedRules.length) {
    return `${name} (metodika MMR ČR)`;
  }
  if (!SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) return name;
  return kind === "recommended"
    ? `${name} (metodika MMR ČR)`
    : item.id === "VYB_01"
      ? `${name} (školský zákon č. 561/2004 Sb., § 179)`
      : `${name} (školský zákon č. 561/2004 Sb., § 178)`;
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
  if (projectEstimate?.type === "supermarket") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const grid = createElement("div", "item-results-grid");
    const addMetric = (label, value, unit) => {
      const metric = createElement("div", "item-result-metric");
      metric.append(createElement("span", "", label));
      metric.append(createElement("strong", "", formatValue(value, unit)));
      grid.append(metric);
    };
    addMetric("Obyvatelé včetně přírůstku", projectEstimate.localPopulation, "obyvatel");
    addMetric("Minimální populační práh (uživatelský scénář)",
      projectEstimate.minimumPopulation, "obyvatel");
    addMetric("Přírůstek obyvatel na prodejnu", projectEstimate.populationPerStore, "obyvatel");
    addMetric("Doporučená sídelní přítomnost scénáře",
      projectEstimate.totalPopulationMeetsThreshold ? "Ano" : "Ne", "");
    addMetric("Kapacita okolí potvrzena uživatelem",
      projectEstimate.capacityVerified ? "Ano" : "Ne", "");
    addMetric("Počet modelových prodejen", projectEstimate.quantity,
      czechCountUnit(projectEstimate.quantity, "prodejna", "prodejny", "prodejen"));
    addMetric("Užitná plocha", projectEstimate.usableArea, "m²");
    addMetric("Celková podlahová plocha", projectEstimate.totalFloorArea, "m²");
    addMetric("Obestavěný prostor", projectEstimate.volume, "m³");
    addMetric("Plocha pozemku", projectEstimate.siteArea, "m²");
    addMetric("Zpevněná plocha", projectEstimate.pavedArea, "m²");
    addMetric("Orientační investice včetně VRN a rezervy", projectEstimate.cost, "Kč");
    metrics.append(grid);
    metrics.append(createElement("p", "item-results-note",
      `Uživatelský scénář supermarketu, nikoli standard MMR ani zákonná povinnost. Cena ${projectEstimate.rateSource === "RTS" ? "vychází z RTS JKSO 801.8." : "je orientační scénářová a není potvrzenou sazbou RTS."}`));
    return metrics;
  }
  if (projectEstimate?.type === "targeted") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const projectMetrics = renderProjectEstimateMetrics(item, kind);
    if (projectMetrics) metrics.append(projectMetrics);
    metrics.append(createElement("p", "item-results-note",
      "Návrhová poptávka, kapacity, výměry a orientační investice jsou dopočtené ze zobrazených scénářových předpokladů. Ověřená okolní nabídka a skutečná dostupnost mohou potřebu snížit; scénář sám neurčuje právní povinnost ani nenahrazuje projekt a rozpočet."));
    return metrics;
  }
  if (item.socialDemographic && projectEstimate?.type === "item") {
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    const grid = createElement("div", "item-results-grid");
    const addMetric = (label, value, unit = "") => {
      const metric = createElement("div", "item-result-metric");
      metric.append(createElement("span", "", label));
      metric.append(createElement("strong", "", value === null || value === undefined
        ? "Nelze vypočítat" : formatValue(value, unit)));
      grid.append(metric);
    };
    const population = numericOrNull(derivedValues().O_prah);
    const coefficient = numericOrNull(getItemEstimateValue(item,
      "SOCIAL_DEMOGRAPHIC_COEFFICIENT"));
    const unroundedCapacity = population === null || coefficient === null ? null
      : population * coefficient;
    addMetric("Současní + noví + spádoví obyvatelé", population, "obyvatel");
    addMetric(`Demografický koeficient (${item.socialDemographic.unitLabel} / obyvatele)`,
      coefficient === null ? null : formatNumber(coefficient, 6));
    addMetric("Demografická potřeba před zaokrouhlením", unroundedCapacity,
      item.socialDemographic.unitLabel);
    addMetric("Návrhová kapacita", projectEstimate.quantity,
      czechCountUnit(projectEstimate.quantity, ...item.socialDemographic.unitForms));
    addMetric("Užitná plocha", projectEstimate.usableArea, "m²");
    addMetric("Celková podlahová plocha", projectEstimate.totalFloorArea, "m²");
    addMetric("Obestavěný prostor", projectEstimate.volume, "m³");
    addMetric("Plocha pozemku", projectEstimate.siteArea, "m²");
    addMetric("Orientační investice", projectEstimate.cost, "Kč");
    metrics.append(grid);
    metrics.append(createElement("p", "item-results-note",
      `Potřeba vychází z celkového počtu obyvatel po záměru včetně spádu a z koeficientu uvedeného přímo ve výpočtu. U odhadových koeficientů nejde o právní kvótu. ${item.socialDemographic.source}`));
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
    if (projectEstimate.isPark) {
      addMetric("Modelový počet parků", projectEstimate.parkCount, "parků");
      addMetric("Doporučená výměra parkové zeleně", projectEstimate.minimumArea, "m²");
      addMetric("Agregovaná parková úprava", projectEstimate.zrn, "Kč");
    } else {
      addMetric("Kvalifikovaná zastavitelná plocha", projectEstimate.qualifyingAreaHa, "ha");
      addMetric(kind === "recommended"
        ? "Doporučená výměra veřejného prostranství"
        : "Zákonné minimum veřejného prostranství", projectEstimate.minimumArea, "m²");
      addMetric("Zpevněná plocha k ocenění", projectEstimate.pavedArea, "m²");
    }
    addMetric("Orientační investice včetně VRN a rezervy", projectEstimate.cost, "Kč");
    metrics.append(grid, createElement("p", "item-results-note", projectEstimate.isPark
      ? "Doporučení MMR se liší od zákonného minima veřejného prostranství: samostatně sleduje parkovou zeleň, její rozsah a pěší dosah. Cena je editovatelný agregovaný odhad."
      : "Započítá se pouze související zastavitelná plocha bydlení, rekreace, občanského vybavení nebo smíšená obytná; pozemní komunikace se do minima nezahrnují. Zkontrolujte skladbu ploch v územním plánu."));
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
      "Jde o souhrnný scénář zdravotnických kapacit, ploch a investic, nikoli o samostatný požadavek na výstavbu. Pro každý druh péče se použije doporučená varianta, pokud existuje, jinak zákonná. Koeficient snižuje souhrnné plochy a investici, nikoli potřebnou kapacitu ani počet ordinací."));
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
  if (kind === "recommended" && MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) {
    const recommendation = getMmrSchoolRecommendation(item);
    if (!recommendation) return metrics;
    metrics.append(createElement("h4", "", RESULTS_HEADING));
    addMetric("Obyvatelé včetně spádu", recommendation.catchmentPopulation, "obyvatel");
    addMetric("Poptávka dle MMR před odečtem okolí", recommendation.totalDemand, getCapacityResultUnit(item));
    addMetric("Ověřená volná kapacita v dosahu", recommendation.capacityVerified
      ? recommendation.appliedCapacity : 0, getCapacityResultUnit(item));
    addMetric("Orientační minimum bodů podle pěšího dosahu",
      recommendation.spatialPointCount, "bodů");
    addMetric("Kapacita pro plošné pokrytí (scénář MMR × minima zařízení)",
      recommendation.spatialCapacityFloor, getCapacityResultUnit(item));
    addMetric("Doporučená kapacita k zajištění", recommendation.requiredCapacity, getCapacityResultUnit(item));
    addMetric("Populační práh jedné standardní jednotky", recommendation.minimumPopulation, "obyvatel");
    addMetric("Minimální kapacita jednoho zařízení při dosaženém prahu", recommendation.siteMinimumCapacity,
      getCapacityResultUnit(item));
    addMetric("Minimální počet nových zařízení podle kapacity / pokrytí",
      recommendation.minimumFacilityCount, "zařízení");
    addMetric("Kapacitní rozpětí jednoho zařízení (aktivní práh)",
      Number.isFinite(recommendation.siteMaximumCapacity)
        ? `${formatNumber(recommendation.siteMinimumCapacity)}–${formatNumber(recommendation.siteMaximumCapacity)}`
        : `${formatNumber(recommendation.siteMinimumCapacity)} a více`,
      "míst");
    if (recommendation.smallSchoolPopulation) {
      addMetric("Prahová hodnota malotřídní školy", recommendation.smallSchoolPopulation, "obyvatel");
      addMetric("Minimální kapacita malotřídní školy", recommendation.smallSchoolCapacity,
        getCapacityResultUnit(item));
    }
    addMetric("Dostupnost dle MMR", recommendation.accessStatus, "");
    if (recommendation.accessRule.density !== undefined) {
      addMetric("Hustota / metodické minimum",
        recommendation.density === null ? null
          : `${formatNumber(recommendation.density)} / ${formatNumber(recommendation.accessRule.density)} obyv./ha`,
        recommendation.densityMeets === null ? "" : recommendation.densityMeets ? "vyhovuje" : "nevyhovuje");
    }
    metrics.append(grid);
    metrics.append(createElement("p", "item-results-note",
      "Doporučená poptávka vychází z MMR poměru obyvatel na dítě/žáka a zahrnuje řešené území po záměru i spád. Prahová hodnota popisuje minimální velikost standardní jednotky, nikoli podmínku pro nulovou potřebu. Volná kapacita se odečte pouze po doložení jejího zdroje a skutečné dostupnosti."));
    return metrics;
  }
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
    } else if ([RESIDENT_PARKING_LAW_ID, RESIDENT_PARKING_RECOMMENDED_ID,
      RESIDENT_PARKING_RESERVE_ID].includes(item.id)) {
      metrics.append(createElement("h4", "", RESULTS_HEADING));
      const isModelRecommendation = item.id === RESIDENT_PARKING_RECOMMENDED_ID;
      const isReserve = item.id === RESIDENT_PARKING_RESERVE_ID;
      addMetric(isModelRecommendation
        ? "Počet stání v historickém srovnávacím scénáři"
        : isReserve ? "Počet stání v samostatném navýšení"
          : "Počet stání v úplné zákonné variantě",
      isModelRecommendation ? recommended?.recommendedStalls ?? "Chybí vstup"
        : isReserve ? recommended?.reserveStalls ?? "Chybí vstup"
          : legal.requiredStalls ?? "Chybí vstup",
      isModelRecommendation || isReserve ? recommended === null ? "" : "stání"
        : legal.requiredStalls === null ? "" : "stání");
      addMetric("Modelová potřeba smíšené zástavby",
        recommended?.modeledStalls ?? "Chybí vstup",
        recommended === null ? "" : "stání");
      if (isModelRecommendation || isReserve) {
        addMetric("Podíl parkovací dům / zpevněná plocha",
          `${formatNumber(getParkingEstimateValue(item, kind, "PARKING_BUILDING_SHARE") * 100)} / ${formatNumber((1 - getParkingEstimateValue(item, kind, "PARKING_BUILDING_SHARE")) * 100)}`,
          "%");
      }
      addMetric("Stání splňující vyhlášku č. 146/2024 Sb.", legal.requiredStalls ?? "Chybí vstup",
        legal.requiredStalls === null ? "" : "stání");
      const parkingMetrics = renderProjectEstimateMetrics(item, kind);
      if (parkingMetrics) {
        [...parkingMetrics.children]
          .filter((metric) => !metric.textContent.startsWith("Počet stání v propočtu")
            && !metric.textContent.startsWith("Podíl parkovací dům / zpevněná plocha"))
          .forEach((metric) => grid.append(metric));
      }
      metrics.append(grid);
      metrics.append(createElement("p", "item-results-note",
        isReserve && recommended !== null
          ? `Samostatné navýšení činí ${formatNumber(recommended.reserveStalls)} stání z ${formatNumber(recommended.selectedBaseStalls)} stání ve zvolené základní variantě. Rezervní položka se započítává místo základní varianty.`
          : recommended === null || legal.requiredStalls === null
          ? "Doplňte počet bytů, HPP nebo ruční zákonnou plochu."
          : `Jde o nadstandardní srovnávací scénář dřívějšího normového režimu (vyhláška č. 268/2009 Sb., § 5 odst. 2; předpis je zrušen). Modelová potřeba je ${formatNumber(recommended.modeledStalls)} stání; výsledek aplikace současně nepoklesne pod kontrolní minimum vyhlášky č. 146/2024 Sb. (${formatNumber(legal.requiredStalls)}), tedy činí ${formatNumber(recommended.recommendedStalls)} stání. Zvolte jej místo povinné varianty; množství se nesčítají.`));
    } else {
      metrics.append(createElement("h4", "", RESULTS_HEADING));
      const selectedBase = recommended?.selectedBaseStalls;
      addMetric("Počet stání ve zvolené základní variantě",
        selectedBase ?? "Nejprve vyberte variantu",
        selectedBase === null || selectedBase === undefined ? "" : "stání");
      addMetric("Procento rezervy", recommended?.reservePercent === undefined
        ? "Chybí vstup" : recommended.reservePercent * 100, "%");
      addMetric("Rezerva k samostatné realizaci", recommended?.reserveStalls ?? "Chybí vstup",
        recommended?.reserveStalls === null || recommended === null ? "" : "stání");
      metrics.append(grid);
      metrics.append(createElement("p", "item-results-note",
        recommended === null
          ? "Doplňte počet bytů, HPP nebo ruční zákonnou plochu."
          : selectedBase === null
            ? "Rezerva zatím není určena. Vyberte nejprve zákonnou, nebo doporučenou parkovací variantu."
            : `Rezerva = zaokrouhlení nahoru (${formatNumber(selectedBase)} × ${formatNumber(recommended.reservePercent * 100)} %) = ${formatNumber(recommended.reserveStalls)} stání. Základem je výhradně zvolená úplná varianta.`));
    }
    return metrics;
  }
  if (isMmrWasteYard(item)) {
    addMetric("Počet obyvatel řešeného území pro standard MMR",
      numericOrNull(derivedValues().O_uzemi_cil), "obyvatel");
  } else if (item.demand !== null) {
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
    const spatialPointRequirement = getSpatialPointRequirement(item, kind);
    addMetric(usesModelCountAsReference
      ? "Počet zařízení / bodů v základním modelu" : "Počet zařízení / bodů N_návrh",
    usesModelCountAsReference ? item.modelProposed
      : Number.isFinite(spatialPointRequirement) ? spatialPointRequirement
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
    const parkingEstimate = getProjectPlanningEstimate(item, kind);
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
    } else if ([RESIDENT_PARKING_LAW_ID, RESIDENT_PARKING_RECOMMENDED_ID].includes(item.id)) {
      const isModelRecommendation = item.id === RESIDENT_PARKING_RECOMMENDED_ID;
      calculation.append(createElement("span", "calculation-kicker",
        isModelRecommendation ? "ÚPLNÁ DOPORUČENÁ VARIANTA" : "ÚPLNÁ ZÁKONNÁ VARIANTA"));
      calculation.append(createElement("code", "calculation-equation",
        recommended === null
          ? "Variantu nelze vypočítat bez počtu jednotek a zákonného základu."
          : isModelRecommendation
            ? `Modelová potřeba = ${formatNumber(recommended.apartmentUnits)} bytů × ${formatNumber(recommended.apartmentRate)} stání/byt + ${formatNumber(recommended.familyHomeUnits)} rodinných domů × ${formatNumber(recommended.familyHomeRate)} stání/dům = ${formatNumber(recommended.modeledStalls)} stání; doporučená kapacita = max(modelová potřeba, zákonné minimum ${formatNumber(legal.requiredStalls)} dle vyhlášky č. 146/2024 Sb.) = ${formatNumber(recommended.recommendedStalls)} stání.`
            : `Zákonná potřeba dle vyhlášky č. 146/2024 Sb. = ${formatNumber(legal.baseStalls)} před zaokrouhlením → ${formatNumber(legal.requiredStalls)} stání.`));
      content.append(calculation);
      content.append(createElement("p", "calculation-explanation",
        recommended === null || legal.requiredStalls === null
          ? "Doplňte počet bytů, HPP nebo ruční zákonnou plochu, aby bylo možné určit obě složky doporučené varianty."
          : isModelRecommendation
          ? `Doporučená úplná varianta = vyšší z modelové potřeby (${formatNumber(recommended?.modeledStalls)}) a zákonného minima (${formatNumber(legal.requiredStalls)}) = ${formatNumber(recommended?.recommendedStalls)} stání. Jde o celou alternativu, nikoli o rozdíl k zákonné položce.`
          : "Jde o úplnou zákonnou alternativu. Zvolte ji místo doporučené varianty; množství se nesčítají."));
    } else {
      calculation.append(createElement("span", "calculation-kicker", "REZERVA NAD RÁMEC ZVOLENÉ VARIANTY"));
      calculation.append(createElement("code", "calculation-equation",
        recommended === null
          ? "Rezervu nelze spočítat bez zákonného minima a modelové potřeby."
          : recommended.selectedBaseStalls === null
            ? "Vyberte jednu základní variantu parkování; rezerva se odvodí pouze z jejího počtu."
            : `Zvolená základní varianta = ${formatNumber(recommended.selectedBaseStalls)} stání; rezerva = zaokrouhlení nahoru (${formatNumber(recommended.selectedBaseStalls)} × ${formatNumber(recommended.reservePercent * 100)} %) = ${formatNumber(recommended.reserveStalls)} stání.`));
      content.append(calculation);
      content.append(createElement("p", "calculation-explanation",
        "Rezerva se realizuje jako samostatná položka navíc k právě zvolené zákonné nebo doporučené variantě. Základní varianty se nesčítají. Procento upravíte tlačítkem „Upravit výpočet“."));
    }
    if (parkingEstimate?.type === "parking") {
      content.append(renderProjectEstimateCalculation(item, kind));
    }
    section.append(content);
    return section;
  }
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  if (item.socialDemographic) {
    const population = numericOrNull(derivedValues().O_prah);
    const coefficient = numericOrNull(getItemEstimateValue(item,
      "SOCIAL_DEMOGRAPHIC_COEFFICIENT"));
    const rawCapacity = population === null || coefficient === null
      ? null : population * coefficient;
    const populationParts = ["O_sidlo", "O_celk", "O_spad"]
      .map((code) => numericOrNull(getPrimary(code)));
    const populationFormula = populationParts.some((value) => value === null)
      ? `O_prah = současní obyvatelé řešeného území + přírůstek záměru + obyvatelé spádu = ${formatValue(population, "obyvatel")}.`
      : `O_prah = ${populationParts.map((value) => formatNumber(value)).join(" + ")} = ${formatNumber(population)} obyvatel.`;
    const socialEquations = [
      populationFormula,
      rawCapacity === null
        ? "Demografická potřeba = nelze vypočítat bez úplných vstupů."
        : `Demografická potřeba = ${formatNumber(population)} obyvatel × ${formatNumber(coefficient, 6)} ${item.socialDemographic.unitLabel} / obyvatele = ${formatNumber(rawCapacity)} ${item.socialDemographic.unitLabel}.`,
      rawCapacity === null
        ? "Návrhová kapacita = nelze vypočítat."
        : `Návrhová kapacita = zaokrouhlení nahoru (${formatNumber(rawCapacity)}) = ${formatNumber(projectEstimate.quantity)} ${czechCountUnit(projectEstimate.quantity, ...item.socialDemographic.unitForms)}.`
    ];
    content.append(renderProjectEstimateCalculation(item, kind, socialEquations));
    content.append(createElement("p", "calculation-explanation",
      `Prah dostupnosti používá stejný součet současných, nových a spádových obyvatel. Výměrové předpoklady jsou orientační, nikoli právní kvóta. ${item.socialDemographic.source}`));
    section.append(content);
    return section;
  }
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
  if (["infrastructure", "infrastructure-reserve"].includes(projectEstimate?.type)) {
    const calculation = renderProjectEstimateCalculation(item, kind);
    if (calculation) content.append(calculation);
    content.append(createElement("p", "calculation-explanation",
      projectEstimate.type === "infrastructure-reserve"
        ? "Rezerva představuje editovatelný procentní přírůstek kapacity a orientační ceny. Při jejím výběru se základní položka této infrastruktury do realizace nezapočítá."
        : "Jde o scénářovou kapacitní a investiční aproximaci. Nejde o samostatný požadavek na výstavbu, vyjádření správce sítě, RTS položku ani projektový rozpočet."));
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
      `Jde o souhrnný návrh zdravotnických kapacit, nikoli samostatný požadavek na výstavbu. Koeficient efektivity ${formatNumber(polyclinicEstimate.discountCoefficient)} snižuje souhrnné plochy a investici kvůli sdílení budovy; počet ordinací ani jejich kapacita se nesnižují.`));
    section.append(content);
    return section;
  }
  if (kind === "recommended" && MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) {
    content.append(renderProjectEstimateCalculation(item, kind));
    section.append(content);
    return section;
  }

  const equations = [];
  const sources = [];
  const comparison = getStandardComparison(item);
  const isStandardDelta = kind === "recommended" && !isCompleteStandardMode()
    && comparison.delta !== null && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id);

  if (item.demand !== null && !isMmrWasteYard(item)
    && !(kind === "recommended" && MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id))) {
    const demandCalculation = formatWorkbookCalculation(
      item.sourceFormulas?.demand, getDemandLabel(item), item.demand, getDemandUnit(item)
    );
    equations.push(demandCalculation || `${getDemandLabel(item)} = ${formatNumber(item.demand)} ${text(getDemandUnit(item), "")}.`);
    const citation = calculationSource("Kapacitní poptávka", item.sourceFormulas?.demand);
    if (citation) sources.push(citation);
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
  if (kind === "recommended" && isMmrWasteYard(item)) {
    const estimate = calculateSpatialEstimate(item);
    if (Number.isFinite(estimate?.count)) {
      equations.push(
        `Počet sběrných dvorů podle standardu MMR = max(${formatNumber(item.modelProposed ?? 0)} v základním modelu, ${formatNumber(estimate.count)} podle územní dostupnosti) = ${formatNumber(getSpatialPointRequirement(item, kind))} areálů.`
      );
      sources.push("TB050MMR01, tabulka 3, tištěná str. 16: typy A/B dochozí vzdálenost 1 500 m; typy C/D jedno zařízení na 5 000 obyvatel. Jde o doporučenou dostupnost, nikoli povinnost postavit nové zařízení.");
    }
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
  if (displayName !== cleanSourceName && item.id !== "VYB_20" && !PARKING_SHARE_ASSUMPTIONS[item.id]) {
    content.append(createElement("p", "original-name-note", `Název ve zdrojovém sešitu: ${cleanSourceName}`));
  }
  const ruleBadges = createElement("div", "rule-badge-row");
  const isCombinedStandard = kind === "recommended" && isCompleteStandardMode()
    && hasStandardComparison(item);
  const applicability = isResidentParkingLegalSource ? "Zákonný výpočet"
    : isCombinedStandard ? "Právní základ + doporučený standard"
      : kind === "legal" ? "Právní požadavek / kontrola"
      : isCompleteStandardMode() ? "Ve zvoleném standardu" : "Volitelné doporučení";
  ruleBadges.append(createElement("span", `applicability-badge${kind === "recommended" && !isCompleteStandardMode() ? " is-optional" : ""}`, applicability));
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

function renderSocialAvailabilityFields(standard, item, kind) {
  const section = createElement("section", "social-availability-check");
  section.append(createElement("strong", "", standard.thresholdText));
  if (item.socialDemographic) {
    const population = numericOrNull(derivedValues().O_prah);
    const coefficient = numericOrNull(getItemEstimateValue(item,
      "SOCIAL_DEMOGRAPHIC_COEFFICIENT"));
    const required = getProjectPlanningEstimate(item, kind)?.quantity;
    const verifiedCapacity = numericOrNull(getCapacityEdit(item).quantity);
    const check = createElement("div", "social-capacity-check");
    check.append(createElement("strong", "", "Demografická potřeba k porovnání s kapacitou okolí"));
    check.append(createElement("p", "calculation-equation",
      population === null || coefficient === null
        || required === null || required === undefined
        ? "Potřebu nelze vypočítat bez populačních vstupů a demografického koeficientu."
        : `${formatNumber(population)} obyvatel × ${formatNumber(coefficient, 6)} ${item.socialDemographic.unitLabel} / obyvatele = ${formatNumber(required)} ${czechCountUnit(required, ...item.socialDemographic.unitForms)} k zajištění.`));
    check.append(createElement("p", "item-results-note",
      verifiedCapacity === null
        ? "Doplňte níže ověřenou volnou kapacitu okolních zařízení; výpočet zatím neurčuje zůstatek."
        : `Zadaná ověřená kapacita okolí: ${formatNumber(verifiedCapacity)} ${item.socialDemographic.unitLabel}.`));
    section.append(check);
  }
  if (standard.details) {
    section.append(createElement("p", "item-results-note", standard.details));
  }
  section.append(createElement("p", "social-standard-applicability",
    getSocialServiceApplicability(standard)));
  const saved = state.socialServiceChecks?.coverage?.[standard.key] || {};
  const statusLabel = createElement("label", "social-service-check-field");
  statusLabel.append(createElement("span", "", "Přítomnost služby v obci"));
  const select = document.createElement("select");
  select.dataset.socialServiceField = "status";
  select.dataset.socialServiceKey = standard.key;
  [
    ["", "Neověřeno"],
    ["present", "Přítomnost zjištěna"],
    ["absent", "Přítomnost nezjištěna"]
  ].forEach(([value, label]) => select.append(new Option(label, value)));
  select.value = saved.status || "";
  statusLabel.append(select);
  const sourceLabel = createElement("label", "social-service-check-field");
  sourceLabel.append(createElement("span", "", "Zdroj / poskytovatel (nepovinné)"));
  const source = document.createElement("input");
  source.type = "text";
  source.value = saved.source || "";
  source.dataset.socialServiceField = "source";
  source.dataset.socialServiceKey = standard.key;
  sourceLabel.append(source);
  section.append(statusLabel, sourceLabel);
  section.append(createElement("p", "calculation-source-note", standard.demographic
    ? `Kapacita používá součet současných, nových a spádových obyvatel. Koeficient je editovatelný přímo ve výpočtu. ${standard.demographic.source} Cena je samostatný orientační odhad, nikoli norma ani rozpočet.`
    : `Kapacita, výměry a cena jsou převzaty jako upravitelný výchozí scénář z položky „${standard.templateName || "katalogové předlohy"}“. Jde o odhad pro plánování, nikoli samostatnou normu sociální služby. Ověření sídelní dostupnosti je samostatné doporučení a samo o sobě neukládá povinnost stavět.`));
  return section;
}

function renderItemCard(item, kind) {
  const isSocialServiceItem = item.category === "SOCIÁLNÍ SLUŽBY A DOSTUPNÉ BYDLENÍ";
  const legalRules = item.legalRules.filter((rule) =>
    !GENERAL_OBLIGATION_RULE_IDS.has(String(rule[RULE_HEADERS.id])) || item.id === "VYB_17B");
  const rules = item.id === "VYB_05B" ? []
    : kind === "legal" ? legalRules
    : kind === "recommended" ? RECOMMENDED_SCENARIO_ITEM_IDS.has(item.id)
        ? [...item.otherRules, ...item.recommendedRules] : item.recommendedRules : item.otherRules;
  const index = categoryIndex(item.category);
  const calc = adjustedResult(item, kind);
  const isSchoolRecommendation = kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
  const isDelta = kind === "recommended" && !isCompleteStandardMode()
    && hasStandardComparison(item) && !isSchoolRecommendation;
  const comparison = isDelta ? getStandardComparison(item) : null;
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  const isTargetedEstimate = projectEstimate?.type === "targeted";
  const isSupermarketEstimate = projectEstimate?.type === "supermarket";
  const housingEstimate = getHousingContributionEstimate(item);
  const isPolyclinicEstimate = projectEstimate?.type === "polyclinic";
  const isInfrastructureEstimate = ["infrastructure", "infrastructure-reserve"]
    .includes(projectEstimate?.type);
  const isPublicSpaceEstimate = projectEstimate?.type === "public-space";
  const isCulturalHallEstimate = item.id === "VYB_10B_1" && projectEstimate?.type === "item";
  const targetedHealthcareRemaining = isTargetedEstimate && isHealthcareDemandItem(item)
    ? projectEstimate.metrics.find((metric) => metric.label.startsWith(
      item.id === "VYB_09A_4"
        ? "Poptávka po gynekologii po odečtu ověřené kapacity"
        : "Poptávka nepokrytá ověřenou kapacitou"))?.value
    : null;
  const isZero = isTargetedEstimate ? isHealthcareDemandItem(item)
    ? targetedHealthcareRemaining === 0 : false
    : isSupermarketEstimate ? projectEstimate.quantity === 0
    : item.socialDemographic ? projectEstimate?.quantity === 0
    : housingEstimate ? housingEstimate.municipalApartments === 0
    : isPolyclinicEstimate ? projectEstimate.clinicCount === 0
    : isInfrastructureEstimate ? false
    : isPublicSpaceEstimate ? false
    : isCulturalHallEstimate ? projectEstimate.quantity === 0
    : isDelta ? comparison.delta === 0 : isRequirementSatisfied(calc);
  const realizationPlan = getRealizationPlan();
  const isSelectedForRealization = realizationPlan.selectedKeys.has(item.key);
  const capacitySatisfied = state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true;
  const standardInactive = kind === "recommended" && isMmrRequirementInactive(item);
  const reserveUnavailable = item.id === RESIDENT_PARKING_RESERVE_ID
    && !getSelectedResidentParkingBase();
  const details = createElement("details",
    `requirement-card category-${index % CATEGORY_COLORS.length} shape-${categoryShape(item.category)}${isZero ? " is-zero" : ""}${isSelectedForRealization ? " is-realization-selected" : ""}${capacitySatisfied ? " is-capacity-satisfied" : ""}${standardInactive ? " is-standard-inactive" : ""}${reserveUnavailable ? " is-parking-reserve-unavailable" : ""}`);
  details.dataset.itemKey = item.key;
  details.dataset.kind = kind;
  const summary = document.createElement("summary");
  const top = createElement("div", "requirement-card-top");
  const realizationButton = createElement("button",
    `category-dot-control${isSelectedForRealization ? " is-selected" : ""}`);
  realizationButton.type = "button";
  realizationButton.dataset.realizationItemKey = item.key;
  const selectionUnavailable = capacitySatisfied || reserveUnavailable;
  realizationButton.disabled = selectionUnavailable;
  realizationButton.setAttribute("aria-pressed", String(isSelectedForRealization));
  realizationButton.setAttribute("aria-label",
    `${isSelectedForRealization ? "Zrušit výběr" : "Vybrat"} k realizaci: ${displayCatalogItemName(item)}`);
  realizationButton.title = isSelectedForRealization ? "Odebrat z realizace" : "Vybrat k realizaci";
  if (isMmrRequirementInactive(item)) {
    realizationButton.title = getInactiveRecommendationText(item);
  } else if (capacitySatisfied) {
    realizationButton.title = "Kapacita okolí je označena jako dostatečná.";
  } else if (reserveUnavailable) {
    realizationButton.title = "Nejprve vyberte zákonnou nebo doporučenou variantu parkování.";
  }
  realizationButton.append(createElement("span", `category-dot shape-${categoryShape(item.category)}`));
  top.append(realizationButton);
  const title = createElement("span", "requirement-title");
  const displayName = kind === "review" ? getReviewAlternativeName(item) : displayRequirementItemName(item, kind);
  title.append(createElement("strong", "", displayName));
  top.append(title);
  const targetedCost = isTargetedEstimate
    ? projectEstimate.metrics.find((metric) => metric.label.includes("investice"))?.value
    : null;
  const summaryQuantityText = item.socialDemographic
    ? `${formatNumber(projectEstimate.quantity)} ${czechCountUnit(projectEstimate.quantity, ...item.socialDemographic.unitForms)}`
    : calc.hasOverride
    ? formatRequirementSummaryAmount(item, calc.required)
    : isTargetedEstimate
    ? getTargetedPlanSummaryText(item, projectEstimate)
    : isSupermarketEstimate
    ? projectEstimate.quantity === null ? "nelze vypočítat"
      : `${formatNumber(projectEstimate.quantity)} ${czechCountUnit(projectEstimate.quantity,
        "modelová prodejna", "modelové prodejny", "modelových prodejen")}`
    : item.id === "VYB_13A"
    ? calc.required === null ? "není vypočteno"
      : `${formatNumber(calc.required)} ${czechCountUnit(calc.required,
        "maloobchodní prodejna", "maloobchodní prodejny", "maloobchodních prodejen")}`
    : item.manualCoverageCheck
    ? "ověřit pokrytí"
    : housingEstimate ? `${formatNumber(housingEstimate.municipalApartments)} bytů pro město`
    : isPolyclinicEstimate ? `${formatNumber(projectEstimate.clinicCount)} ordinací`
    : ["infrastructure", "infrastructure-reserve"].includes(projectEstimate?.type)
    ? `${formatNumber(projectEstimate.quantity)} ${projectEstimate.unit}`
    : isPublicSpaceEstimate ? `${formatNumber(projectEstimate.minimumArea)} m²`
    : isCulturalHallEstimate ? projectEstimate.quantity > 0 ? `${formatNumber(projectEstimate.quantity)} sál` : "prověřit podle populace"
    : isSchoolRecommendation ? getMmrSchoolSummaryText(item)
    : isDelta ? comparison.delta === null ? "navýšení nezadáno"
      : `${formatStandardComparisonAmount(item, comparison.delta)} navíc`
      : reserveUnavailable ? "nejprve vyberte variantu"
      : item.id === "VYB_05B" ? getItemEstimateQuantity(item) > 0
        ? `${formatNumber(getItemEstimateQuantity(item))} ${getItemEstimateQuantity(item) === 1 ? "hala" : "haly"}`
        : "bez samostatné haly"
      : getSummaryRequirementQuantity(calc) === null ? "bez počtu"
        : formatRequirementSummaryAmount(item, getSummaryRequirementQuantity(calc));
  const summaryQuantity = createElement("span", "requirement-quantity",
    formatRequirementSummaryQuantity(item, kind, summaryQuantityText));
  const summaryMetrics = createElement("span", "requirement-summary-metrics");
  summaryMetrics.append(summaryQuantity);
  if (calc.usesDemandBasedCapacity && !isDelta) {
    const capacitySummary = createElement("small", "requirement-capacity", getCapacitySummaryText(item, calc));
    capacitySummary.hidden = !calc.isCapacityVerified;
    summaryMetrics.append(capacitySummary);
  }
  const investmentCost = (item.manualCoverageCheck && !isTargetedEstimate)
    || (isTargetedEstimate && (targetedCost === null || targetedCost === undefined))
    ? null
    : isTargetedEstimate ? targetedCost
      : projectEstimate?.cost !== null && projectEstimate?.cost !== undefined
        ? projectEstimate.cost
        : getDisplayedCost(item, calc, kind);
  summaryMetrics.append(createElement("small", "requirement-investment",
    investmentCost === null || investmentCost === undefined
      ? "bez samostatného ocenění"
      : formatCost(investmentCost)));
  top.append(summaryMetrics);
  top.append(createElement("span", "expand-mark", "＋"));
  summary.append(top);
  details.append(summary);

  const body = createElement("div", "requirement-body");
  body.append(createElement("p", "item-brief-rule", getItemBriefRule(item, kind, rules)));
  if (standardInactive) {
    body.append(createElement("p", "standard-inactive-note",
      item.id === "VYB_17A"
        ? "Doporučení interních metodik policie se na typ D nevztahuje; modelovaná potřeba zůstává uvedena a o zařazení do realizace rozhoduje uživatel."
        : "Mimo práh doporučení MMR pro zvolený typ území nebo počet obyvatel. Modelovaná potřeba zůstává uvedena; o zařazení do realizace rozhoduje uživatel."));
  }
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
  const assumptions = isSocialServiceItem ? null : renderObjectAssumptionControls(item, kind);
  if (assumptions) body.append(assumptions);
  if (isInfrastructureReserveItem(item)) {
    body.append(renderInfrastructureReserveControls(item));
  }
  const availabilitySection = renderCapacityEditor(item, kind);
  if (parkingLegalInputs) availabilitySection.append(parkingLegalInputs);
  if (isDelta || isSchoolRecommendation
    || (kind === "recommended" && isCompleteStandardMode() && hasStandardComparison(item))) {
    availabilitySection.append(renderStandardComparison(item));
  }
  if (item.quantityBasis && !isResidentParkingItem(item) && !isMmrWasteYard(item)) {
    availabilitySection.append(createElement("p", "quantity-basis-note",
      `Počet je převzat z „${MODEL_HEADERS.demand}“ (${formatNumber(item.demand)} ${text(getDemandUnit(item), "")}); hodnota N_navrh v modelu byla ${formatNumber(item.sourceModelQuantity)} a představovala počet zařízení/bodů, nikoli kapacitu uvedenou v jednotce položky.`));
  }
  if (healthEditor) availabilitySection.append(healthEditor);
  if (item.socialAvailabilityStandard) {
    availabilitySection.append(renderSocialAvailabilityFields(item.socialAvailabilityStandard, item, kind));
  }
  if (kind === "recommended") {
    const spatialEstimate = renderSpatialEstimate(item);
    if (spatialEstimate) {
      const spatialDisclosure = createElement("details", "item-section spatial-section");
      addDisclosureSummary(spatialDisclosure, "Prostorový odhad dostupnosti");
      spatialDisclosure.append(spatialEstimate);
      availabilitySection.append(spatialDisclosure);
    }
  }
  if (item.id === "VYB_06A" && kind === "recommended") {
    availabilitySection.append(renderSocialServiceAvailability(item));
  }
  const alternativeNote = renderAlternativeContext(item);
  if (alternativeNote) availabilitySection.append(alternativeNote);
  body.append(availabilitySection);
  if (!isSocialServiceItem
    && (rules.length || item.id === RESIDENT_PARKING_ID && kind === "legal")) {
    body.append(renderRulesDisclosure(item, kind, rules, displayName));
  }
  const transferredWarnings = isSocialServiceItem
    ? null : renderTransferredReviewWarnings(item, rules);
  if (transferredWarnings) body.append(transferredWarnings);
  body.append(createFeedbackButton(
    `${item.category} / ${displayName} [${item.id}] · seznam: ${kind}`
  ));
  details.append(body);
  return details;
}

function renderStandardComparison(item) {
  const saved = state.standardComparisons[item.key] || { legal: "", recommended: "", source: "" };
  const result = getStandardComparison(item);
  const section = createElement("section", "standard-comparison");
  if (SCHOOL_CAPACITY_ITEM_IDS.has(item.id)) {
    const recommendation = getMmrSchoolRecommendation(item);
    section.append(createElement("h4", "", "Doporučený standard dostupnosti MMR"));
    section.append(createElement("p", "standard-comparison-note",
      "Nejde o libovolné procentní navýšení: výpočet používá poměr MMR obyvatel na dítě/žáka, populační práh a pravidlo dostupnosti podle typu území. Údaje o skutečné docházce/VHD a volné kapacitě se zadávají v části ověřeného stavu."));
    section.append(createElement("p", "standard-comparison-result",
      recommendation ? `${formatNumber(recommendation.totalDemand)} ${text(getCapacityResultUnit(item), "míst")} dle MMR; ${formatNumber(recommendation.requiredCapacity)} ${text(getCapacityResultUnit(item), "míst")} k zajištění. Prah jedné standardní jednotky: ${formatNumber(recommendation.minimumPopulation)} obyvatel / ${formatNumber(recommendation.minimumCapacity)} míst.`
        : "Doplňte obyvatelstvo a plochu řešeného území."));
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
    const schoolRecommendation = isSchoolRecommendation ? getMmrSchoolRecommendation(item) : null;
    const isDelta = kind === "recommended" && !isCompleteStandardMode()
      && hasStandardComparison(item) && !isSchoolRecommendation;
    const calc = adjustedResult(item, kind);
    const resultLine = card.querySelector(".standard-comparison-result");
    if (resultLine) {
      resultLine.textContent = isSchoolRecommendation
        ? schoolRecommendation
          ? `${formatNumber(schoolRecommendation.totalDemand)} ${text(getCapacityResultUnit(item), "míst")} dle MMR; ${formatNumber(schoolRecommendation.requiredCapacity)} ${text(getCapacityResultUnit(item), "míst")} k zajištění.`
          : "Doplňte obyvatelstvo a plochu řešeného území."
        : result.delta === null
          ? `Zadejte doporučené navýšení v ${text(getStandardComparisonUnit(item), "jednotce vybavenosti")}.`
          : `Doporučené navýšení ${formatStandardComparisonAmount(item, result.uplift)}; započtená kapacita okolí ${formatStandardComparisonAmount(item, result.appliedCapacity)}; zbývá pokrýt ${formatStandardComparisonAmount(item, result.delta)}; orientační cena navýšení ${formatCost(result.cost)}.`;
    }
    if (!isSchoolRecommendation && !isCompleteStandardMode()
      && !card.querySelector(".requirement-title small")) {
      card.querySelector(".requirement-title")?.append(createElement("small", "", "Rezerva a navýšení kapacity"));
    }
    const quantity = card.querySelector(".requirement-quantity");
    const investment = card.querySelector(".requirement-investment");
    const projectEstimate = getProjectPlanningEstimate(item, kind);
    const displayedInvestment = projectEstimate?.cost ?? getDisplayedCost(item, calc, kind);
    if (isSchoolRecommendation) {
      if (quantity) quantity.textContent = formatRequirementSummaryQuantity(
        item, kind, getMmrSchoolSummaryText(item));
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
  const config = getSpatialEstimateConfig(item);
  if (!config) return null;
  const rule = item.recommendedRules?.find((candidate) => candidate[RULE_HEADERS.id] === config.ruleId);
  if (!rule) return null;
  const territory = getResidentParkingTerritoryType();
  const settlementType = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  let distanceMeters = config.distanceMeters;
  let note = null;
  if (distanceMeters === "school-access-rule") {
    const accessRule = getMmrSchoolAccessRule(item);
    distanceMeters = accessRule.mode === "distance" ? accessRule.distance : null;
    if (distanceMeters === null) note = `Plošný počet nelze odvodit: ${accessRule.summary}`;
  } else if (distanceMeters === "by-settlement-type") {
    if (item.id === "VYB_03C") {
      distanceMeters = settlementType === "D" ? 1000 : 500;
    } else {
      distanceMeters = settlementType === "D" ? null : 300;
      if (distanceMeters === null) {
        note = "Pro typ D metodika požaduje park v každém obytném sídle; počet sídel nelze odvodit z plochy samotné.";
      }
    }
  } else if (distanceMeters === "healthcare-group-1") {
    const localPopulation = numericOrNull(derivedValues().O_uzemi_cil);
    distanceMeters = ["A", "B", "C"].includes(territory) && localPopulation > 2000
      ? 600 : null;
    if (distanceMeters === null) {
      note = ["A", "B", "C"].includes(territory)
        ? "MMR stanoví 600 m a přítomnost v obci pouze při více než 2 000 obyvatelích; při nižším počtu se prostorové minimum z této metodiky neodvozuje."
        : "Pro typ D metodika stanoví dojezdovou, nikoli pevnou pěší vzdálenost; ověřte dostupnost ručně.";
    }
  } else if (distanceMeters === "day-service") {
    distanceMeters = ["A", "B", "C"].includes(territory) ? 600 : null;
    if (distanceMeters === null) note = "Pro typ D metodika stanoví dojezdový standard; plošný počet pěších bodů z něj nelze odvodit.";
  } else if (distanceMeters === "community-center") {
    distanceMeters = ["A", "B"].includes(territory)
      ? getPrimary("KOMPAKT") === "1" ? 400 : 600
      : ["C", "D"].includes(territory) ? 800 : null;
  } else if (distanceMeters === "public-transport-stop") {
    distanceMeters = ["A", "B"].includes(territory)
      ? getPrimary("KOMPAKT") === "1" ? 300 : 500
      : ["C", "D"].includes(territory) ? 700 : null;
    if (["C", "D"].includes(territory)) {
      note = "Pro typy C/D je použit horní okraj doporučeného rozpětí 600–700 m jako konzervativní scénář; ověřte skutečnou pěší trasu.";
    }
  } else if (distanceMeters === "recycling-collection") {
    distanceMeters = ["A", "B", "C"].includes(territory) ? 150 : null;
    if (distanceMeters === null) note = "Pro typ D se ověřuje obsluha každého sídla; počet sídel nelze odvodit pouze z plochy.";
  } else if (distanceMeters === "recycling-yard") {
    if (["A", "B"].includes(territory)) {
      distanceMeters = 1500;
    } else if (["C", "D"].includes(territory)) {
      const population = numericOrNull(derivedValues().O_uzemi_cil);
      if (population === null) {
        return { rule, distanceMeters: null, note: "Doplňte počet obyvatel řešeného území; pro typy C/D se ověřuje 1 zařízení na 5 000 obyvatel." };
      }
      const count = Math.ceil(population / 5000);
      return {
        rule,
        mode: "population",
        count,
        population,
        populationPerFacility: 5000,
        note: "MMR pro typy C/D stanoví 1 zařízení na 5 000 obyvatel, nikoli pevnou pěší vzdálenost."
      };
    } else {
      distanceMeters = null;
    }
  }
  if (distanceMeters === null) {
    return {
      rule,
      distanceMeters: null,
      note: note || "Metodika pro tuto kombinaci nestanoví číselnou pěší vzdálenost; ověřte dostupnost podle skutečné sítě."
    };
  }
  if (!Number.isFinite(distanceMeters)) {
    return {
      rule,
      distanceMeters: null,
      note: note || "Metodika nestanoví číselnou pěší vzdálenost použitelnou pro plošný odhad."
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

function renderSocialServiceAvailability(item) {
  const section = createElement("details", "item-section social-service-availability");
  addDisclosureSummary(section, "Orientační výměry sociálních služeb");
  const content = createElement("div", "social-service-availability-content");
  content.append(createElement("p", "spatial-access-guidance",
    "Potřeby dostupnosti jsou vedené u jednotlivých sociálních služeb v seznamu vybavenosti. Níže uvedené výměry jsou pouze referenční výpočty ÚÚR, nikoli automatická kapacita zařízení."));

  const areaSection = createElement("section", "uur-social-area-reference");
  areaSection.append(createElement("h4", "", "Orientační výměry podle ÚÚR, s. C.4–71 až 72"));
  const projectPopulation = numericOrNull(derivedValues().O_prah);
  if (projectPopulation === null) {
    areaSection.append(createElement("p", "spatial-access-guidance",
      "Doplňte současné obyvatele, přírůstek záměru a obyvatele spádu pro přepočet účelových jednotek."));
  } else {
    areaSection.append(createElement("p", "spatial-access-guidance",
      `Výpočet používá ${formatNumber(projectPopulation)} obyvatel (současný stav + přírůstek záměru + spád). Jde o doporučené optimum ÚÚR, nikoli právní minimum; účelové jednotky mohou vycházet zlomkově.`));
    const table = createElement("div", "uur-social-area-table");
    UUR_SOCIAL_AREA_REFERENCES.forEach((reference) => {
      const units = projectPopulation * reference.unitsPerThousand / 1000;
      const row = createElement("div", "uur-social-area-row");
      row.append(createElement("strong", "", reference.label));
      row.append(createElement("span", "",
        `${formatNumber(projectPopulation)} × ${formatNumber(reference.unitsPerThousand)} / 1 000 = ${formatNumber(units)} ${reference.unitLabel}`));
      row.append(createElement("span", "",
        `${formatNumber(units)} × ${formatNumber(reference.floorAreaPerUnit)} m² = ${formatNumber(units * reference.floorAreaPerUnit)} m² podlažní plochy`));
      row.append(createElement("span", "",
        `${formatNumber(units)} × ${formatNumber(reference.siteAreaPerUnit)} m² = ${formatNumber(units * reference.siteAreaPerUnit)} m² pozemku`));
      table.append(row);
    });
    const halfway = createElement("div", "uur-social-area-row");
    halfway.append(createElement("strong", "", "Dům na půl cesty"));
    halfway.append(createElement("span", "", "ÚÚR uvádí příklad pro 20 klientů."));
    halfway.append(createElement("span", "", "Orientační pozemek: cca 1 500 m²."));
    halfway.append(createElement("span", "", "Normativ podlažní plochy ani poměr na 1 000 obyvatel zde uveden není."));
    table.append(halfway);
    areaSection.append(table);
  }
  areaSection.append(createElement("p", "calculation-source-note",
    "ÚÚR uvádí také orientační výměry pro konkrétní typové příklady, nikoli pro každý druh sociální služby. Zejména u nocleháren a krizové pomoci nelze bez dalšího zaměnit plochu ubytovny či azylového domu za normativ dané služby. Profesionální pečovatelka u denního stacionáře není totéž co klientské místo."));
  content.append(areaSection);
  section.append(content);
  return section;
}

function getSocialDayServiceAccess(item) {
  const territory = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  const edit = state.capacityEdits[item.key] || {};
  if (["A", "B", "C"].includes(territory)) {
    const distance = numericOrNull(edit.accessDistance);
    return {
      status: distance === null
        ? "Skutečná pěší dostupnost neověřena."
        : distance <= 600
          ? `Splňuje standard: ${formatNumber(distance)} m z nejvýše 600 m.`
          : `Nesplňuje standard: ${formatNumber(distance)} m z nejvýše 600 m.`
    };
  }
  if (territory === "D") {
    const minutes = numericOrNull(edit.accessVhdMinutes);
    return {
      status: minutes === null
        ? "Skutečný dojezd neověřen."
        : minutes <= 30
          ? `Splňuje standard: ${formatNumber(minutes)} min z nejvýše 30 min.`
          : `Nesplňuje standard: ${formatNumber(minutes)} min z nejvýše 30 min.`
    };
  }
  return { status: "Doplňte kategorii území A–D." };
}

function updateSocialServiceAvailability(event) {
  const input = event.target.closest("[data-social-service-field][data-social-service-key]");
  if (!input) return;
  const key = input.dataset.socialServiceKey;
  const field = input.dataset.socialServiceField;
  const saved = state.socialServiceChecks?.coverage || {};
  const current = saved[key] || { status: "", source: "" };
  state.socialServiceChecks = {
    ...state.socialServiceChecks,
    coverage: {
      ...saved,
      [key]: { ...current, [field]: input.value }
    }
  };
  saveScenario();
}

function getSocialServiceApplicability(standard) {
  const territory = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  const standardDescription = getSocialAvailabilityStandardDescription(standard);
  if (!territory) {
    return `Doporučený standard: ${standardDescription} Doplňte kategorii území A–D pro vyhodnocení použitelnosti.`;
  }
  if (!standard.territories.includes(territory)) {
    return `Doporučený standard: ${standardDescription} Pro zvolený typ ${territory} se toto doporučení neuplatní.`;
  }
  if (standard.accessStandard) {
    return `Doporučený standard: ${standardDescription} Ověřte dostupnost po skutečné trase.`;
  }
  if (standard.minimumPopulation === undefined) {
    return `Doporučený standard: ${standardDescription}`;
  }
  const population = numericOrNull(derivedValues().O_prah);
  if (population === null) {
    return `Doporučený standard: ${standardDescription} Doplňte současné, nové a spádové obyvatele pro vyhodnocení populačního prahu.`;
  }
  const thresholdStatus = population > standard.minimumPopulation
    ? "Populační práh je splněn."
    : "Populační práh není splněn; modelovaná potřeba zůstává uvedena.";
  return `Doporučený standard: ${standardDescription} ${thresholdStatus}`;
}

function getSocialAvailabilityStandardDescription(standard) {
  const descriptions = {
    "day-services": "Typ A/B/C – skutečná pěší docházka do 600 m; typ D – dojezdová doba do 30 minut.",
    "low-threshold-day-center": "Typ A/B/C, obce > 5 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    "ambulatory-social-care": "Typ A/B/C, obce > 5 000 obyvatel – sídelně-strukturální přítomnost v obci pro nízkoprahové denní centrum, nízkoprahové zařízení pro děti a mládež, týdenní stacionář a intervenční centrum. Kontaktní centrum a sociálně aktivizační služby tabulka 1 v tomto řádku samostatně neuvádí.",
    "night-shelter": "Typ A/B, obce > 30 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    "crisis-services": "Typ A/B, obce > 30 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    "halfway-house": "Typ A/B, obce > 50 000 obyvatel – sídelně-strukturální přítomnost v obci.",
    "care-service-housing": "Typ A – domov pro seniory má doporučenou sídelně-strukturální přítomnost v obci. Dům s pečovatelskou službou je zahrnut v této modelové položce; tabulka 1 pro něj samostatný práh neuvádí.",
    "disability-services": "Typ A/B, obce > 30 000 obyvatel – sídelně-strukturální přítomnost v obci."
  };
  return descriptions[standard.key] || standard.thresholdText;
}

function renderSpatialAccessControls(item, kind) {
  if (kind !== "recommended" || !["VYB_03C", "VYB_06A"].includes(item.id)) return null;
  const territory = String(getPrimary("TYP_UZEMI") ?? "").match(/[A-D]/i)?.[0]?.toUpperCase();
  const capacity = getCapacityEdit(item);
  const wrapper = createElement("section", "spatial-access-check");
  const isSport = item.id === "VYB_03C";
  if (isSport) {
    const targetDistance = territory === "D" ? 1000 : ["A", "B", "C"].includes(territory) ? 500 : null;
    wrapper.append(createElement("p", "spatial-access-guidance",
      targetDistance === null
        ? "Doplňte kategorii území A–D. Standard požaduje skutečnou pěší docházku."
        : `Standard TAČR: skutečná pěší docházka nejvýše ${formatNumber(targetDistance)} m pro typ ${territory}.`));
    if (targetDistance !== null) {
      const field = createTextInput("Skutečná pěší trasa k hřišti",
        capacity.accessDistance ?? "", { id: item.key, field: "accessDistance" }, "m");
      const input = field.querySelector("input");
      input.type = "number";
      input.min = "0";
      input.step = "1";
      input.dataset.kind = "number";
      wrapper.append(field);
      const result = createElement("p", "spatial-access-result");
      result.dataset.spatialAccessResult = "sport";
      const distance = numericOrNull(capacity.accessDistance);
      result.textContent = distance === null
        ? "Skutečnou dostupnost zatím nelze vyhodnotit."
        : distance <= targetDistance
          ? `Splňuje standard (${formatNumber(distance)} m z ${formatNumber(targetDistance)} m).`
          : `Nesplňuje standard (${formatNumber(distance)} m z ${formatNumber(targetDistance)} m).`;
      wrapper.append(result);
    }
    return wrapper;
  }

  const distanceTerritory = ["A", "B", "C"].includes(territory);
  const travelTerritory = territory === "D";
  const row = createElement("div", "spatial-access-controls");
  if (distanceTerritory) {
    row.append(createElement("p", "spatial-access-guidance",
      "Standard TAČR: skutečná pěší docházka nejvýše 600 m pro typy A/B/C."));
    const field = createTextInput("Skutečná pěší trasa k centru/stacionáři",
      capacity.accessDistance ?? "", { id: item.key, field: "accessDistance" }, "m");
    const input = field.querySelector("input");
    input.type = "number";
    input.min = "0";
    input.step = "1";
    input.dataset.kind = "number";
    row.append(field);
  } else if (travelTerritory) {
    row.append(createElement("p", "spatial-access-guidance",
      "Standard TAČR: dojezdová doba nejvýše 30 minut pro typ D."));
    const field = createTextInput("Skutečná dojezdová doba",
      capacity.accessVhdMinutes ?? "", { id: item.key, field: "accessVhdMinutes" }, "min");
    const input = field.querySelector("input");
    input.type = "number";
    input.min = "0";
    input.step = "0.1";
    input.dataset.kind = "number";
    row.append(field);
  } else {
    row.append(createElement("p", "spatial-access-guidance",
      "Doplňte kategorii území A–D pro určení pěšího nebo dojezdového standardu."));
  }
  wrapper.append(row);
  return wrapper;
}

function getSpatialEstimateConfig(item) {
  const entry = SPATIAL_ESTIMATE_CONFIG[item.id];
  const configs = Array.isArray(entry) ? entry : entry ? [entry] : [];
  return configs.find((config) => item.recommendedRules
    ?.some((rule) => rule[RULE_HEADERS.id] === config.ruleId)) ?? null;
}

function hasSpatialPointFloor(item) {
  const config = getSpatialEstimateConfig(item);
  return config?.applyPointFloor ?? SPATIAL_POINT_FLOOR_ITEM_IDS.has(item.id);
}

function isMmrWasteYard(item) {
  return getSpatialEstimateConfig(item)?.ruleId === "TECH-MET-003";
}

function displayShapeName(code) {
  const field = PRIMARY_INPUTS.find((candidate) => candidate.code === "KOD_TVAR");
  const index = field?.options.indexOf(code) ?? -1;
  return index < 0 ? code : field.labels[index];
}

function renderSpatialEstimate(item) {
  const config = getSpatialEstimateConfig(item);
  if (!config) return null;
  const estimate = calculateSpatialEstimate(item);
  const section = createElement("section", "spatial-estimate");
  section.append(createElement("h4", "", "Orientační kontrola prostorové dostupnosti"));
  const source = estimate.rule
    ? [estimate.rule[RULE_HEADERS.source], estimate.rule[RULE_HEADERS.location]].filter(Boolean).join(" · ")
    : config.ruleId;
  if (estimate.count === undefined) {
    section.append(createElement("p", "spatial-estimate-note", estimate.note));
    section.append(createElement("p", "spatial-estimate-source", `Pravidlo: ${source}`));
    return section;
  }
  const result = createElement("p", "spatial-estimate-result");
  const countUnit = estimate.mode === "population"
    ? "zařízení" : czechCountUnit(estimate.count, "bod", "body", "bodů");
  result.append(createElement("strong", "", `${formatNumber(estimate.count)} ${countUnit}`));
  if (estimate.mode !== "population") {
    result.append(createElement("span", "", `Mřížka ${estimate.columns} × ${estimate.rows} bodů`));
  }
  section.append(result);
  if (estimate.mode === "population") {
    section.append(createElement("code", "spatial-estimate-formula",
      `N = ceil(${formatNumber(estimate.population)} obyvatel ÷ ${formatNumber(estimate.populationPerFacility)} obyvatel/zařízení) = ${formatNumber(estimate.count)} zařízení`));
    section.append(createElement("p", "spatial-estimate-note", estimate.note));
  } else {
    const formula = createElement("code", "spatial-estimate-formula",
      `R = ${formatNumber(estimate.distanceMeters)} / ${formatNumber(estimate.detour)} = ${formatNumber(estimate.radius)} m; d = √2 × R × (1 − ${formatNumber(estimate.overlap * 100)} %) = ${formatNumber(estimate.spacing)} m; N = ceil(L/d) × ceil(W/d)`);
    section.append(formula);
    section.append(createElement("p", "spatial-estimate-note",
      `Vychází z plochy ${formatNumber(estimate.areaHa)} ha (${estimate.areaCode}), tvaru ${displayShapeName(estimate.shapeCode)} a dosahu ${formatNumber(estimate.distanceMeters)} m po pěší síti. Kružnice jsou převedeny na vzdušný poloměr koeficientem uliční sítě ${formatNumber(estimate.detour)}; překrytí ${formatNumber(estimate.overlap * 100)} % zkracuje rozteč mřížky, nejde o doslovný podíl překryté plochy kruhů.`));
  }
  section.append(createElement("p", "spatial-estimate-note", estimate.mode === "population"
    ? "Výsledek neodečítá existující zařízení ani ověřenou kapacitu; ověřte obecní či meziobecní službu. Doporučený standard není zákonná povinnost."
    : "Konzervativní odhad pokrytí obalového obdélníku zvoleného tvaru při rovnoměrném rozmístění. Neumisťuje body do skutečné uliční sítě ani neověřuje bariéry, vstupy, existující zařízení či kapacitu; polohy musí potvrdit GIS. Doporučený standard není zákonná povinnost."));
  const spatialFloorIsApplied = hasSpatialPointFloor(item)
    && getSpatialCapacityPerPoint(item) !== null;
  if (!spatialFloorIsApplied
    && !MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)
    && !isHealthcareDemandItem(item)) {
    section.append(createElement("p", "spatial-estimate-note",
      config.pointFloorNote || "Počet bodů se automaticky nepřičítá k návrhovému počtu: pro tuto položku není metodicky podložen automatický převod na realizovaná zařízení."));
  }
  if (spatialFloorIsApplied && CAPACITY_QUANTITY_IDS.has(item.id)) {
    section.append(createElement("p", "spatial-estimate-note",
      `Počet bodů se převádí na kapacitní jednotky poměrem základního modelu (${formatNumber(item.demand)} ${text(getDemandUnit(item), "")} ÷ ${formatNumber(item.modelProposed)} zařízení/bod). Tento přenos zachovává jednotky sešitu; ověřte, zda je kapacita jednoho zařízení v základním modelu vhodná pro místní návrh.`));
  }
  if (source) section.append(createElement("p", "spatial-estimate-source", `Metodický podklad: ${source}`));
  if (spatialFloorIsApplied) {
    const button = createElement("button", "text-button spatial-estimate-action", "Přenést bodový odhad jako návrh");
    button.type = "button";
    button.dataset.spatialEstimate = item.key;
    section.append(button);
  }
  return section;
}

function useSpatialEstimate(event) {
  const button = event.target.closest("[data-spatial-estimate]");
  if (!button) return;
  const item = state.items.find((candidate) => candidate.key === button.dataset.spatialEstimate);
  if (!item) return;
  const estimate = calculateSpatialEstimate(item);
  if (!estimate || !Number.isFinite(estimate.count)) return;
  const pointCount = getSpatialPointRequirement(item) ?? estimate.count;
  const capacityPerPoint = getSpatialCapacityPerPoint(item);
  const estimateQuantity = pointCount * capacityPerPoint;
  if (!Number.isFinite(estimateQuantity)) return;
  const card = button.closest(".requirement-card");
  const quantityInput = card?.querySelector('[data-capacity-field="override"]');
  const note = card?.querySelector('[data-capacity-field="overrideNote"]');
  if (!quantityInput || !note) return;
  quantityInput.value = String(estimateQuantity);
  const outputUnit = CAPACITY_QUANTITY_IDS.has(item.id) ? getDemandUnit(item) : item.unit;
  const estimateCountUnit = estimate.mode === "population"
    ? "zařízení" : czechCountUnit(pointCount, "bod", "body", "bodů");
  const estimateBasis = estimate.mode === "population"
    ? `${formatNumber(estimate.population)} obyvatel při limitu ${formatNumber(estimate.populationPerFacility)} obyvatel/zařízení`
    : `mřížka ${estimate.columns} × ${estimate.rows}, dosah ${estimate.distanceMeters} m, překrytí ${formatNumber(estimate.overlap * 100)} %`;
  note.value = `Plošný odhad (${estimate.rule?.[RULE_HEADERS.id] || "bez metodického pravidla"}): ${pointCount} ${estimateCountUnit}; návrhové množství ${formatNumber(estimateQuantity)} ${text(outputUnit, "zařízení/bodů")} podle kapacity základního modelu; ${estimateBasis}. Nutné potvrdit GIS.`;
  updateCapacityEdit({ target: quantityInput });
  updateCapacityEdit({ target: note });
  showNotice(`Odhad ${pointCount} ${estimateCountUnit} byl uložen jako uživatelský návrh. Ověřte skutečnou dostupnost v GIS.`);
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

function renderMmrSchoolAccessControls(item, kind) {
  if (kind !== "recommended" || !MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) return null;
  const recommendation = getMmrSchoolRecommendation(item);
  if (!recommendation) return null;
  const section = createElement("section", "school-access-check");
  section.append(createElement("p", "capacity-search-guide",
    `Dostupnost dle MMR: ${recommendation.accessRule.summary} ${recommendation.accessRule.density === undefined
      ? "" : `; orientační minimální hrubá hustota ${formatNumber(recommendation.accessRule.density)} obyv./ha.`} Započtení volných míst vyžaduje doložený zdroj, kapacitu i vyhovující dostupnost.`));
  const controls = createElement("div", "capacity-controls");
  const capacity = getCapacityEdit(item);
  const addNumberField = (field, label, value, unit, step = 1) => {
    const wrapper = createTextInput(label, value, { id: item.key, field }, unit);
    const input = wrapper.querySelector("input");
    input.type = "number";
    input.min = "0";
    input.step = String(step);
    controls.append(wrapper);
  };
  if (recommendation.accessRule.mode === "distance") {
    addNumberField("accessDistance", "Skutečná pěší vzdálenost k ověřované škole",
      capacity.accessDistance ?? "", "m");
  } else if (recommendation.accessRule.mode === "vhd") {
    addNumberField("accessVhdMinutes", "Skutečný čas VHD k ověřované škole",
      capacity.accessVhdMinutes ?? "", "min", 0.1);
  } else {
    const wrapper = createElement("label", "capacity-field");
    wrapper.append(createElement("span", "capacity-field-label", "Ruční závěr ověření skutečné dostupnosti"));
    const select = document.createElement("select");
    for (const option of ["Neověřeno", "Ano", "Ne"]) select.append(new Option(option, option));
    select.value = capacity.accessStatus || "Neověřeno";
    select.dataset.recordId = item.key;
    select.dataset.capacityField = "accessStatus";
    wrapper.append(select);
    controls.append(wrapper);
  }
  section.append(controls);
  return section;
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
  const schoolRecommendation = MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)
    && kind === "recommended";
  if (!schoolRecommendation) {
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
  }
  const guidance = createElement("p", "capacity-search-guide", capacitySearchGuide(item));
  editor.append(guidance);
  const spatialAccessControls = renderSpatialAccessControls(item, kind);
  if (spatialAccessControls) editor.append(spatialAccessControls);
  const schoolAccessControls = schoolRecommendation ? null : renderMmrSchoolAccessControls(item, kind);
  if (schoolAccessControls) editor.append(schoolAccessControls);
  if (!schoolRecommendation) {
    controls.append(createTextInput("Zdroj / ověření (nepovinné)", capacity.source, { id: item.key, field: "source" }, "Např. GIS, plán nebo správce"));
    controls.append(createTextInput("Datum (doplní se automaticky)", capacity.date, { id: item.key, field: "date" }, "Vyplní se při zadání kapacity"));
    controls.append(createTextInput("Poznámka k ověření", capacity.note, { id: item.key, field: "note" }, "Rozsah, kapacita a omezení"));
  } else {
    editor.append(createElement("p", "capacity-search-guide",
      "Počet míst a plošné pokrytí se navrhuje automaticky z přírůstku obyvatel a prahů MMR. Potvrďte zde, zda skutečná kapacita a dostupnost v okolí tuto potřebu pokrývá."));
  }
  const parkingDemand = renderParkingDemand(item);
  if (parkingDemand) editor.append(parkingDemand);

  const override = state.overrides[item.key] || { quantity: "", note: "", author: "", history: [] };
  const overrideWrap = createTextInput("Uživatelsky upravený návrh", override.quantity, { id: item.key, field: "override", kind: "number" }, "Ponechat hodnotu modelu");
  const overrideInput = overrideWrap.querySelector("input");
  overrideInput.type = "number";
  overrideInput.min = "0";
  overrideInput.step = "any";
  overrideInput.setAttribute("aria-label", `Uživatelsky upravený návrh pro ${displayItemName(item)}`);
  controls.append(overrideWrap);
  controls.append(createTextInput("Autor přepisu (nepovinné)", override.author, { id: item.key, field: "overrideAuthor" }, "Volitelné"));
  controls.append(createTextInput("Důvod / poznámka k úpravě návrhu (nepovinné)", override.note, { id: item.key, field: "overrideNote" }, "Volitelné"));
  editor.append(createElement("p", "capacity-caution capacity-audit-summary",
    getUserOverrideAuditSummary(override, getCapacityUnit(item))));
  editor.append(controls);

  const result = createElement("div", "capacity-result");
  result.dataset.resultFor = item.key;
  updateCapacityResultText(result, item, calc, kind);
  editor.append(result);
  const caution = createElement("p", "capacity-caution",
    "Kapacitu i uživatelský návrh lze zadat bez zdroje, poznámky nebo autora; všechny doplňující údaje jsou dobrovolné. Datum se nastaví automaticky při zadání kapacity. Původní zákonný či metodický výpočet zůstává zachován jako srovnávací základ. U škol a veřejných prostranství se porovnává v kapacitní jednotce (žáci/děti nebo m²), nikoli v počtu budov/souborů.");
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
    "Zadejte procento navíc z právě zvolené zákonné nebo doporučené varianty. Rezerva se přičítá jako samostatná položka."));
  const fields = createElement("div", "secondary-fields project-estimate-inputs");
  const field = {
    code: "PARKING_RESERVE_PERCENT",
    label: "Rezerva nad rámec zvolené základní varianty",
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
  const searchRow = createElement("div", "availability-search-row");
  const link = createElement("a", "text-button availability-search-link",
    "Hledat dostupnost vybavenosti v okolí ↗");
  link.href = availabilitySearchUrl(item);
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  searchRow.append(link);
  searchRow.append(createElement("p", "item-results-note availability-search-note",
    "Odkaz otevře webové vyhledávání podle názvu vybavenosti a lokality. Výsledky nejsou samy o sobě ověřením kapacity; zaznamenejte konkrétní zařízení a ověřený zdroj."));
  section.append(searchRow);
  const key = item.key || item.id;
  const satisfied = state.capacityEdits[key]?.satisfiesNearbyCapacity === true;
  section.append(renderCapacitySatisfactionButton(key, satisfied));
  return section;
}

function renderCapacitySatisfactionButton(itemKey, satisfied) {
  const button = createElement("button",
    `capacity-satisfaction-toggle${satisfied ? " is-confirmed" : ""}`,
    satisfied ? "Splňuje kapacitu v okolí — potvrzeno" : "Splňuje kapacitu v okolí");
  button.type = "button";
  button.dataset.capacitySatisfactionItemKey = itemKey;
  button.setAttribute("aria-pressed", String(Boolean(satisfied)));
  return button;
}

function updateCapacityResultText(target, item, calc, kind) {
  target.replaceChildren();
  const line = createElement("div", "result-line");
  const projectEstimate = getProjectPlanningEstimate(item, kind);
  const isSchoolRecommendation = kind === "recommended" && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
  const isDelta = kind === "recommended" && !isCompleteStandardMode()
    && hasStandardComparison(item) && !isSchoolRecommendation;
  const comparison = isDelta ? getStandardComparison(item) : null;
  if (projectEstimate?.type === "targeted" && isHealthcareDemandItem(item)) {
    const metricValue = (prefix) => projectEstimate.metrics.find((metric) =>
      metric.label.startsWith(prefix))?.value;
    const isSpecialistGroup = ["VYB_09B_1", "VYB_09B_2"].includes(item.id);
    const demandLabel = item.id === "VYB_09A_4"
      ? "Poptávka po gynekologii po odečtu ověřené kapacity"
      : "Poptávka nepokrytá ověřenou kapacitou";
    const requiredFte = metricValue("Požadovaná modelová potřeba")
      ?? metricValue("Indikativní počet lékařských úvazků");
    const clinics = metricValue("Modelový počet ordinací")
      ?? metricValue("Modelový počet odborných ambulancí")
      ?? metricValue("Modelový počet ambulancí");
    const doctorUnit = isSpecialistGroup ? "ambulancí" : "lékařských úvazků";
    line.append(createElement("span", "", `Modelová potřeba: ${formatValue(requiredFte, doctorUnit)}`));
    line.append(createElement("span", "", `Ověřená volná kapacita: ${formatValue(metricValue("Ověřená volná kapacita v okolí") ?? metricValue("Ověřená volná kapacita"), doctorUnit)}`));
    line.append(createElement("strong", "", `Zbývající modelová potřeba: ${formatValue(metricValue(demandLabel), doctorUnit)}`));
    line.append(createElement("strong", "", `${formatNumber(clinics)} ${isSpecialistGroup
      ? czechCountUnit(clinics, "ambulance", "ambulance", "ambulancí")
      : czechCountUnit(clinics, "ordinace", "ordinace", "ordinací")}`));
  } else if (projectEstimate?.type === "supermarket") {
    line.append(createElement("strong", "", projectEstimate.quantity === null
      ? "Počet modelových prodejen nelze vypočítat."
      : `Modelová potřeba: ${formatNumber(projectEstimate.quantity)} ${czechCountUnit(
        projectEstimate.quantity, "prodejna", "prodejny", "prodejen")}`));
  } else if (item.id === "VYB_13A") {
    line.append(createElement("strong", "", calc.required === null
      ? "Počet maloobchodních prodejen nelze vypočítat."
      : `Modelový počet: ${formatNumber(calc.required)} ${czechCountUnit(
        calc.required, "prodejna", "prodejny", "prodejen")}`));
  } else if (isSchoolRecommendation) {
    line.append(createElement("strong", "", `Doporučená kapacita: ${getMmrSchoolSummaryText(item)}`));
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

function getBrnoNumber(key) {
  const value = state.brnoInputs[key];
  if (value === undefined || value === null || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function getBrnoUseHpp(kind, use) {
  const extra = getBrnoNumber(`${kind}-extra-area-${use.key}`) ?? 0;
  if (use.key === "housing") {
    const area = Math.max(0, numericOrNull(derivedValues().HPP_byty) ?? 0);
    return {
      area: area + extra,
      calculatedArea: area,
      extraArea: extra,
      hasData: area > 0 || state.brnoInputs[`${kind}-extra-area-${use.key}`] !== undefined,
      facilities: area > 0 ? [{ name: "Navržené bydlení", area }] : [],
      unestimated: []
    };
  }
  const itemIds = BRNO_HPP_ITEM_IDS[use.key] || [];
  const selectedKeys = getRealizationPlan().selectedKeys;
  let candidates = state.items.filter((item) =>
    selectedKeys.has(item.key) && itemIds.includes(item.id)
    && !isInfrastructureItem(item) && !REALIZATION_NON_BUILDING_IDS.has(item.id)
    && !REALIZATION_OUTDOOR_IDS.has(item.id));
  if (candidates.some((item) => item.id === "VYB_20")) {
    candidates = candidates.filter((item) =>
      !["VYB_09A_4", "VYB_09B_1", "VYB_09B_2", "VYB_16", "VYB_17", "VYB_18", "VYB_19"]
        .includes(item.id));
  }
  const facilities = [];
  const unestimated = [];
  for (const item of candidates) {
    if (item.id === "VYB_19" && !isHealthcareItem(item)) continue;
    const realizationKind = getRealizationExportKind(item);
    const estimate = getProjectPlanningEstimate(item, realizationKind);
    const area = numericOrNull(estimate?.totalFloorArea)
      ?? numericOrNull(estimate?.grossTotalFloorArea)
      ?? numericOrNull(estimate?.facilityPlan?.totalFloorArea)
      ?? numericOrNull(getRealizedFacilityMetrics(item, realizationKind).totalFloorArea);
    if (area === null) {
      unestimated.push(displayCatalogItemName(item));
      continue;
    }
    facilities.push({ name: displayCatalogItemName(item), area: Math.max(0, area) });
  }
  const calculatedArea = facilities.reduce((sum, facility) => sum + facility.area, 0);
  const extraKey = `${kind}-extra-area-${use.key}`;
  return {
    area: calculatedArea + extra,
    calculatedArea,
    extraArea: extra,
    hasData: candidates.length > 0 || state.brnoInputs[extraKey] !== undefined,
    facilities,
    unestimated
  };
}

function getBrnoParkingShare(type) {
  const stored = getBrnoNumber(`${type}-cost-building-share`);
  if (stored !== null) return Math.min(1, stored);
  const parking = state.items.find((item) => item.id === RESIDENT_PARKING_ID);
  return parking
    ? numericOrNull(getParkingEstimateValue(parking, "legal", "PARKING_BUILDING_SHARE")) ?? 0.2
    : 0.2;
}

function getBrnoCostEstimate(item, quantity) {
  if (!Number.isFinite(quantity) || quantity < 0) return null;
  const surcharge = getBrnoNumber(`${item.type}-cost-surcharge`) ?? 0.08;
  if (item.type === "trees") {
    const areaPerSpace = getBrnoNumber("trees-cost-area-per-space") ?? 25;
    const ratePerArea = getBrnoNumber("trees-cost-rate-per-area") ?? 500;
    const area = quantity * areaPerSpace;
    const zrn = area * ratePerArea;
    return { quantity, buildingQuantity: 0, surfaceQuantity: quantity,
      areaPerUnit: areaPerSpace, pavedArea: area, volume: 0, zrn,
      surcharge, cost: Math.round(zrn * (1 + surcharge)), ratePerArea };
  }
  const parking = state.items.find((candidate) => candidate.id === RESIDENT_PARKING_ID);
  const parkingValue = (code, fallback) => parking
    ? numericOrNull(getParkingEstimateValue(parking, "legal", code)) ?? fallback : fallback;
  const buildingShare = getBrnoParkingShare(item.type);
  const volumePerUnit = getBrnoNumber(`${item.type}-cost-volume-per-unit`)
    ?? parkingValue("PARKING_VOLUME_PER_STALL_M3", 18);
  const buildingRate = getBrnoNumber(`${item.type}-cost-building-rate`)
    ?? parkingValue("PARKING_BUILDING_RATE_M3", 8995);
  const surfaceRate = getBrnoNumber(`${item.type}-cost-surface-rate`)
    ?? parkingValue("PARKING_SURFACE_RATE_M3", 7500);
  const surfaceAreaPerUnit = getBrnoNumber(`${item.type}-cost-surface-area-per-unit`)
    ?? parkingValue("PARKING_SURFACE_AREA_PER_STALL_M2", 25);
  const buildingQuantity = quantity * buildingShare;
  const surfaceQuantity = quantity - buildingQuantity;
  const buildingVolume = buildingQuantity * volumePerUnit;
  const surfaceVolume = surfaceQuantity * volumePerUnit;
  const pavedArea = surfaceQuantity * surfaceAreaPerUnit;
  const zrn = buildingVolume * buildingRate + surfaceVolume * surfaceRate;
  return {
    quantity, buildingShare, buildingQuantity, surfaceQuantity, volumePerUnit,
    buildingVolume, surfaceVolume, buildingRate, surfaceRate,
    surfaceAreaPerUnit, pavedArea, zrn, surcharge,
    cost: Math.round(zrn * (1 + surcharge))
  };
}

function getBrnoProposedQuantity(item, calculatedQuantity) {
  const override = numericOrNull(state.overrides[item.key]?.quantity);
  return override !== null && override >= 0 ? override : calculatedQuantity;
}

function getBrnoCarEstimate() {
  const zone = state.brnoInputs["cars-zone"];
  const zoneRates = {
    "00": { otherMinimum: 0, housingMinimum: 0, otherMaximum: 15 },
    "01": { otherMinimum: 15, housingMinimum: 50, otherMaximum: null },
    "02": { otherMinimum: 35, housingMinimum: 75, otherMaximum: null },
    "03": { otherMinimum: 75, housingMinimum: 100, otherMaximum: null }
  }[zone];
  if (!zoneRates) return null;
  let housingBase = 0;
  let otherBoundBase = 0;
  let visitorBase = 0;
  let hasEnteredArea = false;
  const calculations = [];
  for (const use of BRNO_CAR_USE_RATES) {
    const hpp = getBrnoUseHpp("cars", use);
    if (hpp.unestimated.length) {
      return { incomplete: `Nelze určit HPP realizace: ${hpp.unestimated.join(", ")}.` };
    }
    hasEnteredArea ||= hpp.hasData;
    const area = hpp.area;
    if (area === 0) continue;
    const rate = use.rate ?? getBrnoNumber(`cars-rate-${use.key}`);
    if (!rate || (use.key === "production" && (rate < 200 || rate > 800))) {
      return { incomplete: "U výroby zadejte ukazatel HPP na stání v intervalu 200–800 m²." };
    }
    let baseCount = area / rate;
    if (use.dwellings) {
      const dwellingCount = numericOrNull(derivedValues().BJ_pocet) ?? 0;
      baseCount = Math.min(baseCount, dwellingCount * 2);
      housingBase += baseCount;
    } else {
      otherBoundBase += baseCount * use.bound / 100;
      visitorBase += baseCount * use.visitor / 100;
    }
    calculations.push({ use, hpp: area, rate, baseCount, bound: baseCount * use.bound / 100,
      visitor: baseCount * use.visitor / 100, facilities: hpp.facilities,
      extraArea: hpp.extraArea });
  }
  if (!hasEnteredArea) return null;
  const housingMinimum = Math.ceil(housingBase * zoneRates.housingMinimum / 100);
  const otherBoundMinimum = Math.ceil(otherBoundBase * zoneRates.otherMinimum / 100);
  const visitorMinimum = Math.floor(visitorBase * zoneRates.otherMinimum / 100 + 0.5);
  const otherBoundMaximum = zoneRates.otherMaximum === null
    ? null : Math.ceil(otherBoundBase * zoneRates.otherMaximum / 100);
  const visitorMaximum = zoneRates.otherMaximum === null
    ? null : Math.floor(visitorBase * zoneRates.otherMaximum / 100 + 0.5);
  const estimate = {
    zone,
    housingMinimum,
    otherBoundMinimum,
    visitorMinimum,
    minimum: housingMinimum + otherBoundMinimum + visitorMinimum,
    otherBoundMaximum,
    visitorMaximum,
    maxSubset: otherBoundMaximum === null ? null : otherBoundMaximum + visitorMaximum,
    calculations,
    housingUnits: numericOrNull(derivedValues().BJ_pocet) ?? 0
  };
  const carItem = BRNO_REQUIREMENTS.find((item) => item.type === "cars");
  estimate.proposedQuantity = getBrnoProposedQuantity(carItem, estimate.minimum);
  estimate.cost = getBrnoCostEstimate(carItem, estimate.proposedQuantity);
  estimate.surfaceStalls = estimate.cost?.surfaceQuantity ?? 0;
  return estimate;
}

function getBrnoBikeEstimate() {
  let boundBase = 0;
  let visitorBase = 0;
  let housingBoundBase = 0;
  let housingVisitorBase = 0;
  let hasEnteredArea = false;
  const calculations = [];
  for (const use of BRNO_BIKE_USE_RATES) {
    const hpp = getBrnoUseHpp("bikes", use);
    if (hpp.unestimated.length) {
      return { incomplete: `Nelze určit HPP realizace: ${hpp.unestimated.join(", ")}.` };
    }
    hasEnteredArea ||= hpp.hasData;
    const area = hpp.area;
    if (area === 0) continue;
    const baseCount = area / use.rate;
    boundBase += baseCount * use.bound / 100;
    visitorBase += baseCount * use.visitor / 100;
    if (use.key === "housing") {
      housingBoundBase += baseCount * use.bound / 100;
      housingVisitorBase += baseCount * use.visitor / 100;
    }
    calculations.push({ use, hpp: area, rate: use.rate, baseCount,
      bound: baseCount * use.bound / 100, visitor: baseCount * use.visitor / 100,
      facilities: hpp.facilities, extraArea: hpp.extraArea });
  }
  const dwellingCondition = state.brnoInputs["bikes-dwelling-condition"] === "true";
  if (!hasEnteredArea && !dwellingCondition) return null;
  let bound = Math.ceil(boundBase);
  const visitor = Math.floor(visitorBase + 0.5);
  let dwellingIncrement = 0;
  let dwellingMinimum = null;
  if (dwellingCondition) {
    dwellingMinimum = numericOrNull(derivedValues().BJ_pocet) ?? 0;
    dwellingMinimum *= 2;
    const housingBaseTotal = Math.ceil(housingBoundBase)
      + Math.floor(housingVisitorBase + 0.5);
    dwellingIncrement = Math.max(0, dwellingMinimum - housingBaseTotal);
    bound += dwellingIncrement;
  }
  const total = bound + visitor;
  const estimate = {
    bound,
    visitor,
    total,
    oversized: total >= 50 ? Math.floor(total * 0.02 + 0.5) : 0,
    dwellingMinimum,
    dwellingIncrement,
    calculations,
    housingUnits: numericOrNull(derivedValues().BJ_pocet) ?? 0
  };
  const bikeItem = BRNO_REQUIREMENTS.find((item) => item.type === "bikes");
  estimate.proposedQuantity = getBrnoProposedQuantity(bikeItem, estimate.total);
  estimate.cost = getBrnoCostEstimate(bikeItem, estimate.proposedQuantity);
  return estimate;
}

function getBrnoRequirementEstimate(item) {
  if (item.type === "cars") return getBrnoCarEstimate();
  if (item.type === "bikes") return getBrnoBikeEstimate();
  if (item.type === "trees") {
    const enteredSurfaceStalls = numericOrNull(state.brnoInputs["trees-surface-stalls"]);
    if (enteredSurfaceStalls !== null) {
      const estimate = { stalls: enteredSurfaceStalls, spaces: Math.ceil(enteredSurfaceStalls / 8),
        manuallyEntered: true };
      estimate.cost = getBrnoCostEstimate(item, estimate.spaces);
      return estimate;
    }
    const parking = getBrnoCarEstimate();
    if (!parking) return null;
    if (parking.incomplete) return { incomplete: `Nejprve doplňte výpočet automobilových stání: ${parking.incomplete}` };
    const stalls = parking.surfaceStalls;
    const estimate = { stalls, spaces: Math.ceil(stalls / 8) };
    estimate.cost = getBrnoCostEstimate(item, estimate.spaces);
    return estimate;
  }
  return null;
}

function createBrnoNumberInput(label, key, unit, step = "any") {
  const wrapper = createElement("label", "bsp-project-input");
  wrapper.append(createElement("span", "", `${label}${unit ? ` (${unit})` : ""}`));
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.step = step;
  input.value = state.brnoInputs[key] ?? "";
  input.dataset.bspField = key;
  wrapper.append(input);
  return wrapper;
}

function createBrnoUseInputs(kind, uses) {
  const details = createElement("details", "bsp-use-details");
  details.append(createElement("summary", "",
    `HPP podle účelu užívání (${uses.length} ukazatelů)`));
  const rows = createElement("div", "bsp-use-inputs");
  for (const use of uses) {
    const row = createElement("div",
      `bsp-use-input-row${use.rate === null ? " bsp-use-input-row-variable" : ""}`);
    row.dataset.bspUseKey = use.key;
    row.dataset.bspUseKind = kind;
    const description = createElement("span", "");
    description.append(document.createTextNode(use.label));
    description.append(createElement("small", "",
      `${use.rate ?? "200–800"} m² HPP/stání · vázaná ${use.bound}% · návštěvnická ${use.visitor}%`));
    const hpp = getBrnoUseHpp(kind, use);
    const hppValue = createElement("strong", "bsp-use-hpp-value",
      `${formatNumber(hpp.calculatedArea)} m² HPP z realizací${hpp.extraArea
        ? ` + ${formatNumber(hpp.extraArea)} m² doplněno` : ""}`);
    row.append(description, hppValue);
    if (use.rate === null) {
      const rateInput = createBrnoNumberInput("Ukazatel HPP na 1 stání",
        `${kind}-rate-${use.key}`, "m²", "1");
      const rateControl = rateInput.querySelector("input");
      rateControl.min = "200";
      rateControl.max = "800";
      row.append(rateInput);
    }
    row.append(createBrnoNumberInput("Doplnit HPP mimo vybrané realizace",
      `${kind}-extra-area-${use.key}`, "m²"));
    rows.append(row);
  }
  details.append(rows);
  return details;
}

function createBrnoPercentInput(label, key, value) {
  const field = createBrnoNumberInput(label, key, "%", "1");
  const input = field.querySelector("input");
  input.min = "0";
  input.max = "100";
  input.value = formatNumber(value * 100);
  input.dataset.bspFraction = "true";
  return field;
}

function getBrnoParkingCostDefaults(type) {
  const parking = state.items.find((item) => item.id === RESIDENT_PARKING_ID);
  const readParking = (code, fallback) => parking
    ? numericOrNull(getParkingEstimateValue(parking, "legal", code)) ?? fallback : fallback;
  return {
    buildingShare: getBrnoParkingShare(type),
    volumePerUnit: readParking("PARKING_VOLUME_PER_STALL_M3", 18),
    buildingRate: readParking("PARKING_BUILDING_RATE_M3", 8995),
    surfaceRate: readParking("PARKING_SURFACE_RATE_M3", 7500),
    surfaceAreaPerUnit: readParking("PARKING_SURFACE_AREA_PER_STALL_M2", 25),
    surcharge: 0.08
  };
}

function createBrnoCostControls(item) {
  const details = createElement("details", "bsp-cost-details");
  details.append(createElement("summary", "", "Orientační výměry a cena — upravitelné předpoklady"));
  const content = createElement("div", "bsp-cost-inputs");
  if (item.type === "trees") {
    content.append(
      createBrnoNumberInput("Plocha prostoru pro růst jednoho stromu", "trees-cost-area-per-space", "m²", "1"),
      createBrnoNumberInput("Modelová úprava plochy", "trees-cost-rate-per-area", "Kč/m²", "100"),
      createBrnoPercentInput("VRN a rezerva", "trees-cost-surcharge",
        getBrnoNumber("trees-cost-surcharge") ?? 0.08)
    );
    content.append(createElement("p", "item-results-note",
      "Cena je editovatelný scénářový odhad zelené plochy mimo BSP; výchozí sazba 500 Kč/m² není cenovou sazbou předpisu."));
  } else {
    const defaults = getBrnoParkingCostDefaults(item.type);
    content.append(
      createBrnoPercentInput("Podíl kapacity v parkovacím domě / krytém objektu",
        `${item.type}-cost-building-share`, defaults.buildingShare),
      createBrnoNumberInput("Obestavěný prostor na jednotku", `${item.type}-cost-volume-per-unit`, "m³/jednotku", "1"),
      createBrnoNumberInput("Modelová sazba objektu", `${item.type}-cost-building-rate`, "Kč/m³ OP", "100"),
      createBrnoNumberInput("Modelová sazba venkovní plochy", `${item.type}-cost-surface-rate`, "Kč/m³ ekv.", "100"),
      createBrnoNumberInput("Zpevněná plocha na venkovní jednotku",
        `${item.type}-cost-surface-area-per-unit`, "m²/jednotku", "1"),
      createBrnoPercentInput("VRN a rezerva", `${item.type}-cost-surcharge`, defaults.surcharge)
    );
    content.append(createElement("p", "item-results-note",
      "Cena odděluje krytou kapacitu a venkovní zpevněnou plochu stejně jako model rezidentského parkování. Všechny sazby a výměry jsou editovatelné scénářové odhady mimo BSP; objemový ekvivalent venkovní plochy není položkovou cenou RTS."));
  }
  details.append(content);
  return details;
}

function renderBrnoRequirementResult(card, item) {
  const estimate = getBrnoRequirementEstimate(item);
  const summary = card.querySelector("[data-bsp-summary]");
  const result = card.querySelector("[data-bsp-result]");
  result.replaceChildren();
  result.append(createElement("h4", "", RESULTS_HEADING));
  const grid = createElement("div", "item-results-grid");
  result.append(grid);
  const addLine = (label, value) => {
    const metric = createElement("div", "item-result-metric");
    metric.append(createElement("span", "", label), createElement("strong", "", value));
    grid.append(metric);
  };
  const investment = card.querySelector("[data-bsp-investment]");
  if (investment) investment.textContent = estimate?.cost
    ? formatCost(estimate.cost.cost) : "";
  const addHppBreakdown = () => {
    if (!estimate.calculations?.length) return;
    const heading = createElement("h4", "", "HPP a kapacita podle účelu");
    result.append(heading);
    const table = createElement("div", "bsp-calculation-rows");
    estimate.calculations.forEach(({ use, hpp, rate, baseCount, bound, visitor, facilities, extraArea }) => {
      const row = createElement("div", "bsp-calculation-row");
      const detail = [
        `${formatNumber(hpp)} m² HPP ÷ ${formatNumber(rate)} m²/stání = ${formatNumber(baseCount, 2)} základních stání`,
        `vázaná část ${formatNumber(bound, 2)} · návštěvnická ${formatNumber(visitor, 2)}`
      ];
      const sources = facilities.map((facility) =>
        `${facility.name}: ${formatNumber(facility.area)} m²`).join("; ");
      if (sources) detail.push(sources);
      if (extraArea) detail.push(`ručně doplněno mimo výběr: ${formatNumber(extraArea)} m²`);
      row.append(createElement("strong", "", use.label),
        createElement("span", "", detail.join(" · ")));
      table.append(row);
    });
    result.append(table);
  };
  if (!estimate) {
    summary.textContent = "Čeká na výpočtové vstupy";
    result.append(createElement("p", "",
      item.type === "trees"
        ? "Zadejte počet venkovních automobilových stání přímo, nebo vyberte zónu a doplňte HPP pro automatický výpočet."
        : "Po výběru realizovaných položek se jejich vypočtená HPP automaticky přiřadí k účelům užívání."));
    return;
  }
  if (estimate.incomplete) {
    summary.textContent = "Výpočet neúplný";
    result.append(createElement("p", "", estimate.incomplete));
    return;
  }
  if (item.type === "cars") {
    summary.textContent = `${formatValue(estimate.minimum, "stání")} minimum${estimate.proposedQuantity !== estimate.minimum
      ? ` · návrh ${formatValue(estimate.proposedQuantity, "stání")}` : ""} · ${formatCost(estimate.cost?.cost)}`;
    addLine("Minimální počet stání", formatValue(estimate.minimum, "stání"));
    addLine("Vázaná stání bydlení", formatValue(estimate.housingMinimum, "stání"));
    addLine("Vázaná stání ostatní", formatValue(estimate.otherBoundMinimum, "stání"));
    addLine("Návštěvnická stání", formatValue(estimate.visitorMinimum, "stání"));
    result.append(createElement("p", "",
      estimate.maxSubset === null
        ? "Příloha č. 3 pro tuto zónu maximum nestanoví."
        : `Zóna A — dílčí maxima: ostatní vázaná ${formatValue(estimate.otherBoundMaximum, "stání")}, návštěvnická ${formatValue(estimate.visitorMaximum, "stání")} (součet ${formatValue(estimate.maxSubset, "stání")}). Předpis nestanoví maximum vázaných stání pro bydlení.`));
    addHppBreakdown();
  } else if (item.type === "bikes") {
    summary.textContent = `${formatValue(estimate.total, "míst")} celkem${estimate.proposedQuantity !== estimate.total
      ? ` · návrh ${formatValue(estimate.proposedQuantity, "míst")}` : ""} · ${formatCost(estimate.cost?.cost)}`;
    addLine("Vázaná místa", formatValue(estimate.bound, "míst"));
    addLine("Návštěvnická místa", formatValue(estimate.visitor, "míst"));
    if (estimate.dwellingMinimum !== null) {
      addLine("Minimum 2 místa na byt",
        `${formatValue(estimate.dwellingMinimum, "míst")} (navýšení ${formatValue(estimate.dwellingIncrement, "míst")})`);
    }
    if (estimate.oversized) {
      addLine("Nadrozměrná místa (započtená do celku)",
        formatValue(estimate.oversized, "míst"));
    }
    addHppBreakdown();
  } else if (item.type === "trees") {
    summary.textContent = `${formatValue(estimate.spaces, "stromových prostorů")} · ${formatCost(estimate.cost?.cost)}`;
    addLine("Venkovní parkovací stání", formatValue(estimate.stalls, "stání"));
    addLine("Prostory pro růst stromů",
      `${formatNumber(estimate.stalls)} ÷ 8${estimate.manuallyEntered ? " (ručně zadaná venkovní stání)" : " (převzato z výpočtu automobilových stání)"}, zaokrouhleno nahoru = ${formatValue(estimate.spaces, "prostorů")}`);
  }
  if (estimate.cost) {
    const cost = estimate.cost;
    if (item.type === "trees") {
      addLine("Orientační plocha zeleně", formatValue(cost.pavedArea, "m²"));
      addLine("Cena", `${formatCost(cost.zrn)} ZRN + ${formatNumber(cost.surcharge * 100)} % VRN = ${formatCost(cost.cost)}`);
    } else {
      if (cost.quantity !== (item.type === "cars" ? estimate.minimum : estimate.total)) {
        addLine("Kapacita použitá pro cenový odhad",
          formatValue(cost.quantity, item.type === "cars" ? "stání" : "míst"));
      }
      addLine("Parkovací dům / krytá kapacita",
        `${formatValue(cost.buildingQuantity, "jednotek")} · ${formatValue(cost.buildingVolume, "m³")}`);
      addLine("Venkovní zpevněná plocha",
        `${formatValue(cost.surfaceQuantity, "jednotek")} · ${formatValue(cost.pavedArea, "m²")}`);
      addLine("Orientační investice",
        `${formatCost(cost.zrn)} ZRN + ${formatNumber(cost.surcharge * 100)} % VRN = ${formatCost(cost.cost)}`);
    }
  }
  result.append(createElement("p", "item-results-note",
    "HPP se přebírá z propočtů právě realizovaných položek; dodatečnou HPP lze zadat pouze pro objekty mimo katalog. Cena je pracovní odhad, nikoli sazba BSP."));
}

function createBrnoRequirementCard(item) {
  const parkingCategory = "DOPRAVA, PARKOVÁNÍ A BEZPEČNOST";
  const card = createElement("details",
    `requirement-card bsp-calculation-card category-${categoryIndex(parkingCategory) % CATEGORY_COLORS.length} shape-${categoryShape(parkingCategory)}`);
  const selected = Number(state.realizationSelections[item.key]) === 1;
  card.dataset.bspItem = item.key;
  card.dataset.itemKey = item.key;
  card.classList.toggle("is-realization-selected", selected);
  card.classList.toggle("is-capacity-satisfied",
    state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true);
  const summary = createElement("summary", "");
  const top = createElement("div", "requirement-card-top");
  const select = createElement("button",
    `category-dot-control${selected ? " is-selected" : ""}`);
  select.type = "button";
  select.dataset.realizationItemKey = item.key;
  select.setAttribute("aria-pressed", String(selected));
  select.setAttribute("aria-label", `${selected ? "Odebrat výběr" : "Vybrat"}: ${item.label}`);
  select.append(createElement("span", `category-dot shape-${categoryShape("BRNĚNSKÁ VYHLÁŠKA")}`, ""));
  const title = createElement("span", "requirement-title");
  title.append(createElement("strong", "", item.label));
  const metrics = createElement("span", "requirement-summary-metrics");
  const estimateSummary = createElement("span", "requirement-quantity", "");
  estimateSummary.dataset.bspSummary = "";
  metrics.append(estimateSummary);
  const investment = createElement("small", "requirement-investment", "");
  investment.dataset.bspInvestment = "";
  metrics.append(investment);
  top.append(select, title, metrics, createElement("span", "expand-mark", "＋"));
  summary.append(top);
  card.append(summary);
  const body = createElement("div", "requirement-body");
  body.append(createElement("p", "item-brief-rule", item.description));
  const result = createElement("section", "item-results bsp-result");
  result.dataset.bspResult = "";
  body.append(result);

  const calculation = createElement("details", "item-section calculation-section");
  addDisclosureSummary(calculation, "Výpočet / jak jsme k výsledku došli");
  const calculationContent = createElement("div", "calculation-content");
  const inputs = createElement("details", "bsp-calculation-inputs project-estimate-controls");
  addDisclosureSummary(inputs, "Upravit vstupní předpoklady a ocenění");
  const fields = createElement("div", "bsp-project-fields");
  if (item.type === "cars") {
    const zone = createElement("label", "bsp-project-input");
    zone.append(createElement("span", "", "Zóna podle mapy přílohy č. 3"));
    const selectZone = document.createElement("select");
    selectZone.dataset.bspField = "cars-zone";
    const blankOption = document.createElement("option");
    blankOption.value = "";
    blankOption.textContent = "Vyberte zónu";
    selectZone.append(blankOption);
    for (const value of ["00", "01", "02", "03"]) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = `Zóna ${String.fromCharCode(65 + Number(value))}`;
      selectZone.append(option);
    }
    selectZone.value = state.brnoInputs["cars-zone"] ?? "";
    zone.append(selectZone);
    fields.append(zone, createElement("p", "item-results-note",
      `Počet bytů ${formatNumber(derivedValues().BJ_pocet ?? 0)} a HPP bydlení přebírá aplikace z odvozených ukazatelů projektu.`));
    inputs.append(fields, createBrnoUseInputs("cars", BRNO_CAR_USE_RATES));
    inputs.append(createElement("p", "calculation-source-note",
      "Výpočet používá účely a podíly podle přílohy č. 2 a zónová minima přílohy č. 3 (§ 25). Zónu ověřte podle polohy a mapy předpisu; územní plán může stanovit jiný podíl. Specifické účely, vyhrazená stání a přechodná ustanovení vyžadují samostatnou kontrolu."));
  } else if (item.type === "bikes") {
    inputs.append(createBrnoUseInputs("bikes", BRNO_BIKE_USE_RATES));
    const conditions = createElement("label", "realization-filter");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.bspField = "bikes-dwelling-condition";
    checkbox.checked = state.brnoInputs["bikes-dwelling-condition"] === "true";
    conditions.append(checkbox, document.createTextNode(
      "Potvrzuji podmínky přílohy č. 4 odst. 7 pro minimum 2 místa na byt"));
    inputs.append(fields, conditions);
    inputs.append(createElement("p", "item-results-note",
      `Ověřovaný počet bytů ${formatNumber(derivedValues().BJ_pocet ?? 0)} se přebírá z odvozených ukazatelů projektu.`));
    inputs.append(createElement("p", "calculation-source-note",
      "Výpočet používá HPP účelů a podíly vázaných/návštěvnických míst podle § 27 a přílohy č. 4. Předpokládá novou stavbu; u změn dokončených staveb ověřte zvláštní pravidla. Vázaná místa mají být zastřešena; běžné místo má nejméně 0,45 × 2 m, nadrozměrné 2,75 × 0,85 m."));
  } else if (item.type === "trees") {
    fields.append(createBrnoNumberInput("Počet venkovních automobilových stání (volitelný přímý vstup)",
      "trees-surface-stalls", "stání", "1"));
    inputs.append(fields);
    inputs.append(createElement("p", "calculation-source-note",
      "Počet lze zadat přímo, nebo se převezme z výpočtu automobilových stání po odečtení podílu kapacity v parkovacím domě. Požadavek upravuje § 26 odst. 7; umístění stromů mimo parkoviště ověřte v projektu."));
  }
  calculationContent.append(inputs);
  calculation.append(calculationContent);
  body.append(calculation, createBrnoCostControls(item), renderBrnoAvailability(item));
  body.append(createFeedbackButton(`Brněnská vyhláška / ${item.label} [${item.key}]`));
  card.append(body);
  renderBrnoRequirementResult(card, item);
  return card;
}

function renderBrnoAvailability(item) {
  const section = createElement("details", "item-section capacity-section bsp-capacity-check");
  addDisclosureSummary(section, "Ověřený stav v okolí a uživatelský návrh");
  section.append(renderAvailabilitySearch({ ...item, name: item.label }));
  const editor = createElement("div", "capacity-editor");
  editor.append(createElement("p", "capacity-search-guide",
    "Ověřte existující kapacitu a dostupnost; označení dostatečné kapacity slouží jako uživatelská kontrola a nemění samo o sobě výpočet zákonného minima."));
  const saved = state.capacityEdits[item.key] || {};
  const controls = createElement("div", "capacity-controls");
  const quantity = createTextInput("Celková ověřená kapacita v okolí (stání / místa)",
    saved.quantity, { id: item.key, field: "quantity", kind: "number" }, "Prázdné = neověřeno");
  const quantityInput = quantity.querySelector("input");
  quantityInput.type = "number";
  quantityInput.min = "0";
  quantityInput.step = "1";
  controls.append(quantity);
  controls.append(createTextInput("Zdroj / ověření (nepovinné)", saved.source,
    { id: item.key, field: "source" }, "Např. GIS, parkovací dům nebo správce"));
  controls.append(createTextInput("Datum (doplní se automaticky)", saved.date,
    { id: item.key, field: "date" }));
  controls.append(createTextInput("Poznámka k ověření (nepovinné)", saved.note,
    { id: item.key, field: "note" }, "Rozsah a omezení"));
  const override = state.overrides[item.key] || {};
  const proposal = createTextInput("Uživatelsky upravený návrh (nepovinné)",
    override.quantity, { id: item.key, field: "override", kind: "number" }, "Ponechat výsledek výpočtu");
  const proposalInput = proposal.querySelector("input");
  proposalInput.type = "number";
  proposalInput.min = "0";
  proposalInput.step = "1";
  controls.append(proposal);
  controls.append(createTextInput("Poznámka k úpravě návrhu (nepovinné)", override.note,
    { id: item.key, field: "overrideNote" }, "Volitelné"));
  editor.append(controls);
  const result = createElement("div", "capacity-result");
  result.dataset.brnoCapacityResult = item.key;
  result.textContent = getBrnoCapacityResultText(item, saved, override);
  editor.append(result);
  section.append(editor);
  return section;
}

function getBrnoCapacityResultText(item, capacity, override) {
  const knownCapacity = numericOrNull(capacity.quantity);
  const proposedQuantity = numericOrNull(override.quantity);
  const observed = knownCapacity === null
    ? "Ověřená kapacita okolí není zadána."
    : `Zaznamenaná ověřená kapacita okolí: ${formatValue(knownCapacity, item.type === "bikes" ? "míst" : "stání")}.`;
  const proposal = proposedQuantity === null
    ? "Výpočetní minimum zůstává beze změny."
    : `Uživatelský návrh ${formatValue(proposedQuantity, item.type === "bikes" ? "míst" : "stání")} se použije pro orientační cenu; zákonné minimum se tím nepřepisuje.`;
  return `${observed} ${proposal}`;
}

function renderBrnoRequirementList() {
  if (!state.items.length) return;
  const selectedKeys = getRealizationPlan().selectedKeys;
  const items = state.hideUnrealizedItems
    ? BRNO_REQUIREMENTS.filter((item) => selectedKeys.has(item.key))
    : BRNO_REQUIREMENTS;
  elements.brnoCount.textContent = String(items.length);
  elements.brnoList.replaceChildren();
  if (!items.length) {
    const empty = createElement("div", "empty-list",
      "V tomto seznamu nejsou vybrané položky brněnských předpisů.");
    empty.dataset.realizationEmpty = "";
    elements.brnoList.append(empty);
  } else {
    items.forEach((item) => elements.brnoList.append(createBrnoRequirementCard(item)));
  }
  const controls = [...elements.brnoList.querySelectorAll("button[data-realization-item-key]")]
    .filter((control) => !control.disabled);
  const selected = controls.length > 0
    && controls.every((control) => control.getAttribute("aria-pressed") === "true");
  for (const button of elements.requirementColumns.querySelectorAll(
    'button[data-realization-section-toggle="brno"]'
  )) {
    button.setAttribute("aria-pressed", String(selected));
    button.classList.toggle("is-selected", selected);
  }
}

function updateBrnoRequirementInputs(event) {
  const field = event.target.closest("[data-bsp-field]");
  if (!field) return;
  const key = field.dataset.bspField;
  const parsedValue = field.type === "checkbox" ? String(field.checked)
    : field.type === "number" && field.value !== "" ? Number(field.value)
      : field.value;
  if (typeof parsedValue === "number"
    && (!Number.isFinite(parsedValue) || parsedValue < Number(field.min || 0)
      || field.max !== "" && parsedValue > Number(field.max))) {
    field.setCustomValidity("Hodnota je mimo povolený rozsah.");
    if (event.type === "change") field.reportValidity();
    return;
  }
  field.setCustomValidity?.("");
  state.brnoInputs[key] = field.dataset.bspFraction === "true"
    && typeof parsedValue === "number" ? parsedValue / 100 : parsedValue;
  saveScenario();
  for (const item of BRNO_REQUIREMENTS) {
    const card = [...elements.brnoList.querySelectorAll("[data-bsp-item]")]
      .find((candidate) => candidate.dataset.bspItem === item.key);
    if (!card) continue;
    renderBrnoRequirementResult(card, item);
    const hppRows = card.querySelectorAll(".bsp-use-input-row");
    for (const row of hppRows) {
      const use = [...BRNO_CAR_USE_RATES, ...BRNO_BIKE_USE_RATES]
        .find((candidate) => candidate.key === row.dataset.bspUseKey);
      if (!use) continue;
      const hpp = getBrnoUseHpp(row.dataset.bspUseKind, use);
      const output = row.querySelector(".bsp-use-hpp-value");
      if (output) output.textContent = `${formatNumber(hpp.calculatedArea)} m² HPP z realizací${hpp.extraArea
        ? ` + ${formatNumber(hpp.extraArea)} m² doplněno` : ""}`;
    }
  }
  updateScenarioStatus();
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
  renderRequirementList(elements.recommendedList, "recommended", matching);
  renderRequirementList(elements.reviewList, "review", matching);
  renderRequirementList(elements.reserveList, "reserve", matching);
  renderRequirementList(elements.infrastructureList, "infrastructure", matching);
  renderBrnoRequirementList();
  for (const [name, list] of [
    ["legal", elements.legalList],
    ["infrastructure", elements.infrastructureList],
    ["recommended", elements.recommendedList],
    ["review", elements.reviewList],
    ["reserve", elements.reserveList],
    ["brno", elements.brnoList]
  ]) {
    const collapsed = Boolean(state.collapsedLists[name]);
    list.hidden = collapsed;
    const button = elements.requirementColumns.querySelector(`[data-collapse-list="${name}"]`);
    if (!button) continue;
    button.setAttribute("aria-expanded", String(!collapsed));
    const sectionName = button.closest(".requirement-column").querySelector("h3").textContent;
    button.setAttribute("aria-label", `${collapsed ? "Rozbalit" : "Sbalit"} ${sectionName}`);
  }
  for (const button of elements.requirementColumns.querySelectorAll(
    "button[data-realization-section-toggle]"
  )) {
    const column = button.closest(".requirement-column");
    const controls = [...column.querySelectorAll(
      ".requirement-list button[data-realization-item-key]"
    )].filter((control) => !control.disabled);
    const selected = controls.length > 0
      && controls.every((control) => control.getAttribute("aria-pressed") === "true");
    button.setAttribute("aria-pressed", String(selected));
    button.classList.toggle("is-selected", selected);
    const sectionName = column.querySelector("h3").textContent;
    button.setAttribute("aria-label",
      `${selected ? "Odebrat výběr" : "Vybrat všechny položky"}: ${sectionName}`);
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
      [
        ["uživatelsky upravený návrh", override.quantity],
        ["autor přepisu", override.author],
        ["výchozí hodnota modelu", override.baselineQuantity],
        ["čas poslední změny", override.updatedAt]
      ]);
    for (const entry of Array.isArray(override.history) ? override.history : []) {
      add(location, "historie uživatelského přepisu", entry.reason, [
        ["autor", entry.author],
        ["čas změny", entry.changedAt],
        ["předchozí množství", entry.previousQuantity],
        ["nové množství", entry.quantity]
      ]);
    }
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
  const location = document.querySelector("#feedback-location").value.trim()
    || getCurrentFeedbackLocation() || "Všechny skupiny vybavenosti";
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

function getCurrentFeedbackLocation() {
  const selectedItem = state.items.find((item) => item.key === state.selectedItemKey);
  if (selectedItem) {
    return `${selectedItem.category} / ${displayItemName(selectedItem)} [${selectedItem.id}]`;
  }
  return state.selectedCategory ? `Tematická skupina: ${state.selectedCategory}` : "";
}

function createFeedbackButton(location) {
  const button = createElement("button", "text-button item-feedback-button",
    "Nahlásit chybu u této položky");
  button.type = "button";
  button.dataset.feedbackContext = location;
  return button;
}

let feedbackReportObjectUrl = "";

function clearPreparedFeedbackReport() {
  if (feedbackReportObjectUrl) URL.revokeObjectURL(feedbackReportObjectUrl);
  feedbackReportObjectUrl = "";
  const downloadLink = document.querySelector("#feedback-download-link");
  downloadLink.hidden = true;
  downloadLink.removeAttribute("href");
  downloadLink.removeAttribute("download");
  const emailLink = document.querySelector("#feedback-email-link");
  emailLink.hidden = true;
  emailLink.removeAttribute("href");
}

function openFeedbackReport(location = getCurrentFeedbackLocation()) {
  const report = document.querySelector("#feedback-report");
  report.open = true;
  document.querySelector("#feedback-location").value = location;
  document.querySelector("#feedback-status").textContent = "";
  clearPreparedFeedbackReport();
  report.scrollIntoView({ behavior: "auto", block: "start" });
  document.querySelector("#feedback-errors").focus({ preventScroll: true });
}

const USER_GUIDE_STEPS = [
  {
    title: "Načtěte pracovní sešit",
    target: "#dropzone",
    text: "Tady načtete soubor Excel (.xlsx, .xls nebo .xlsm), který obsahuje katalog vybavenosti, pravidla a výchozí vstupy. Lze jej také přetáhnout na oblast Importu. Po načtení zkontrolujte název sešitu a případná upozornění; pokud jsou data už načtená, pokračujte.",
    note: "Začněte výběrem pracovního sešitu, který odpovídá vašemu projektu."
  },
  {
    title: "Vyplňte parametry projektu",
    target: ".project-input-group-pinned",
    text: "V horní části postupně vyplňte známé údaje o území, obyvatelích, bytech, podlažnosti a HPP. Jednotky a popisy jsou uvedené u jednotlivých polí. Změna vstupu přepočítá navázané potřeby a výsledky; prázdné volitelné pole nechte prázdné, pokud hodnotu neznáte.",
    note: "Zvýrazněná oblast obsahuje hlavní vstupy projektu."
  },
  {
    title: "Zkontrolujte odvozené ukazatele",
    target: ".derived-details > summary",
    text: "Rozbalte tuto část pro automaticky odvozené hodnoty a případné volitelné přepisy. Ruční přepis použijte jen tehdy, když máte lepší podklad; přepočítané ukazatele se promítnou do položek, které na ně navazují.",
    note: "Zkontrolujte, zda odvozené hodnoty odpovídají vašim předpokladům."
  },
  {
    title: "Najděte potřebnou vybavenost",
    target: "#search-input",
    text: "Hledejte podle názvu, kódu nebo zdroje. Tematické skupiny pod vyhledáváním lze rozbalit a sbalit. Kliknutím na skupinu nebo konkrétní vybavenost omezíte pravý výstup; opětovným kliknutím filtr zrušíte.",
    note: "Vyhledávání a seznam tematických skupin jsou vlevo."
  },
  {
    title: "Vyberte položky k realizaci",
    target: "#category-list",
    text: "Zaškrtněte políčko u každé položky, kterou chcete zahrnout do záměru. Výběr označuje váš projektový scénář, nikoli automaticky právní povinnost. Infrastrukturu, rezervy a návazné vybavenosti zkontrolujte a vyberte zvlášť. Přepínač „Skrýt nerealizované“ později omezí pravý seznam na vybrané položky.",
    note: "Výběr lze kdykoli změnit; položky nejsou předvybrané za vás."
  },
  {
    title: "Projděte výstup a standardy",
    target: ".requirement-column-heading",
    text: "Vpravo jsou výsledky rozdělené do zákonných/povinných položek, doporučení MMR, nadstandardu, rezerv, infrastruktury a Brněnské vyhlášky. Porovnejte, co je povinnost, doporučení nebo pouze modelový scénář. Souhrn karty ukazuje potřebné množství a případnou orientační cenu.",
    note: "Výstup interpretujte podle konkrétního typu pravidla i podmínek území."
  },
  {
    title: "Upravte výpočet položky",
    target: ".requirement-card .calculation-section > summary",
    text: "Otevřete „Výpočet / jak jsme k výsledku došli“. Upravitelné vstupy, koeficienty a předpoklady měňte podle doložených údajů; výsledek, výměry i cena se podle nich přepočítají. U Brněnské vyhlášky jsou vstupy parkování a část „Orientační výměry a cena“ přímo v příslušné kartě. Zdroje a jednotky vždy zkontrolujte.",
    note: "Průvodce otevřel ukázkovou kartu a její výpočet; změny můžete rovnou vyzkoušet."
  },
  {
    title: "Ověřte kapacitu v okolí",
    target: ".requirement-card .capacity-section > summary",
    text: "V části ověření lze hledat okolní zařízení, zaznamenat doloženou volnou kapacitu a připsat zdroj, datum nebo vlastní návrh. Výsledky hledání ověřte samostatně. Prázdné pole znamená neověřeno a samotné nalezení zařízení nepotvrzuje jeho dostupnou kapacitu.",
    note: "Doplňte pouze údaje, které jste schopni ověřit."
  },
  {
    title: "Exportujte seznam realizací",
    target: "#realized-export",
    text: "Po výběru a kontrole položek použijte „Export realizací (Excel)“. Export obsahuje přehled vybraných položek a jejich vypočtené hodnoty. Před exportem zkontrolujte, že jsou zahrnuté potřebné infrastruktury a varianty a že nejsou omylem skryté nebo vynechané.",
    note: "Toto tlačítko exportuje seznam vybraných realizací."
  },
  {
    title: "Uložte přenositelný projekt",
    target: "#excel-export",
    text: "„Export Excel“ uloží přenositelný projektový sešit včetně úprav a scénáře. Uložte si jej jako pracovní kopii; při příštím použití jej znovu načtěte. „Obnovit hodnoty ze sešitu“ vrátí výchozí hodnoty a smaže místní úpravy scénáře.",
    note: "Export realizací je užší přehled; tento export uchovává projekt a jeho úpravy."
  },
  {
    title: "Nahlaste chybu nebo návrh",
    target: "#open-feedback-report",
    text: "Tlačítko v horní liště otevře obecný formulář. U každé položky je také její vlastní tlačítko „Nahlásit chybu u této položky“, které předvyplní místo. Popište očekávaný výsledek, co se stalo a kroky k zopakování. Připravte a stáhněte report .txt, potom otevřete e-mail, přiložte soubor a zprávu sami odešlete.",
    note: "Aplikace report ani e-mail nikam neodesílá automaticky."
  }
];

let userGuideStepIndex = 0;
let userGuideHighlight = null;
let userGuideDisclosureState = new Map();

function showUserGuideStep(index) {
  userGuideHighlight?.classList.remove("guide-highlight");
  userGuideHighlight = null;
  userGuideStepIndex = Math.max(0, Math.min(index, USER_GUIDE_STEPS.length - 1));
  const step = USER_GUIDE_STEPS[userGuideStepIndex];
  const guide = document.querySelector("#user-guide-tour");
  document.querySelector("#guide-step-title").textContent = step.title;
  document.querySelector("#guide-step-text").textContent = step.text;
  document.querySelector("#guide-target-note").textContent = step.note;
  document.querySelector("#guide-step-count").textContent =
    `Krok ${userGuideStepIndex + 1} z ${USER_GUIDE_STEPS.length}`;
  document.querySelector("#guide-progress-bar").style.width =
    `${((userGuideStepIndex + 1) / USER_GUIDE_STEPS.length) * 100}%`;
  document.querySelector("#guide-back").disabled = userGuideStepIndex === 0;
  document.querySelector("#guide-next").textContent =
    userGuideStepIndex === USER_GUIDE_STEPS.length - 1 ? "Dokončit" : "Další";

  const target = [...document.querySelectorAll(step.target)]
    .find((candidate) => candidate.getClientRects().length
      || candidate.closest(".requirement-card")?.getClientRects().length);
  if (target) {
    for (let ancestor = target; ancestor; ancestor = ancestor.parentElement) {
      if (ancestor instanceof HTMLDetailsElement && !ancestor.open) {
        const key = getUserGuideDisclosureKey(ancestor);
        if (!userGuideDisclosureState.has(key)) userGuideDisclosureState.set(key, ancestor.open);
        ancestor.open = true;
      }
    }
    userGuideHighlight = target;
    target.classList.add("guide-highlight");
    target.scrollIntoView({ behavior: "auto", block: "center", inline: "nearest" });
  }
  positionUserGuideWindow(target);
  guide.querySelector(".guide-window").focus({ preventScroll: true });
}

function getUserGuideDisclosureKey(details) {
  if (details.id) return `id:${details.id}`;
  if (details.dataset.inputGroup) return `input-group:${details.dataset.inputGroup}`;
  const card = details.closest(".requirement-card");
  if (card) {
    if (details === card) return `card:${card.dataset.kind}:${card.dataset.itemKey}:root`;
    const index = [...card.querySelectorAll("details")].indexOf(details);
    return `card:${card.dataset.kind}:${card.dataset.itemKey}:section:${index}`;
  }
  return `details:${[...details.classList].sort().join(".")}`;
}

function positionUserGuideWindow(target) {
  const windowElement = document.querySelector(".guide-window");
  if (!target) {
    windowElement.dataset.placement = "bottom-right";
    return;
  }
  const targetRect = target.getBoundingClientRect();
  const windowRect = windowElement.getBoundingClientRect();
  const gap = 16;
  const placements = [
    { name: "top-left", left: gap, top: gap },
    { name: "top-right", left: window.innerWidth - windowRect.width - gap, top: gap },
    { name: "bottom-left", left: gap, top: window.innerHeight - windowRect.height - gap },
    { name: "bottom-right", left: window.innerWidth - windowRect.width - gap, top: window.innerHeight - windowRect.height - gap }
  ];
  const best = placements.reduce((currentBest, placement) => {
    const overlapWidth = Math.max(0,
      Math.min(placement.left + windowRect.width, targetRect.right)
      - Math.max(placement.left, targetRect.left));
    const overlapHeight = Math.max(0,
      Math.min(placement.top + windowRect.height, targetRect.bottom)
      - Math.max(placement.top, targetRect.top));
    const overlapArea = overlapWidth * overlapHeight;
    return overlapArea < currentBest.overlapArea
      ? { ...placement, overlapArea } : currentBest;
  }, { name: "bottom-right", overlapArea: Infinity });
  windowElement.dataset.placement = best.name;
}

function startUserGuide() {
  userGuideDisclosureState = new Map();
  document.querySelector("#user-guide-tour").hidden = false;
  showUserGuideStep(0);
}

function finishUserGuide() {
  userGuideHighlight?.classList.remove("guide-highlight");
  userGuideHighlight = null;
  for (const disclosure of document.querySelectorAll("details")) {
    const key = getUserGuideDisclosureKey(disclosure);
    if (userGuideDisclosureState.has(key)) {
      disclosure.open = userGuideDisclosureState.get(key);
    }
  }
  userGuideDisclosureState.clear();
  document.querySelector("#user-guide-tour").hidden = true;
  document.querySelector("#open-user-guide").focus({ preventScroll: true });
}

function handleUserGuideKeydown(event) {
  if (document.querySelector("#user-guide-tour").hidden) return;
  if (event.key === "Escape") {
    event.preventDefault();
    finishUserGuide();
  } else if (event.key === "ArrowRight") {
    event.preventDefault();
    if (userGuideStepIndex === USER_GUIDE_STEPS.length - 1) finishUserGuide();
    else showUserGuideStep(userGuideStepIndex + 1);
  } else if (event.key === "ArrowLeft" && userGuideStepIndex > 0) {
    event.preventDefault();
    showUserGuideStep(userGuideStepIndex - 1);
  }
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
  clearPreparedFeedbackReport();
  const downloadLink = document.querySelector("#feedback-download-link");
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
  try {
    const file = new Blob([report], { type: "text/plain;charset=utf-8" });
    feedbackReportObjectUrl = URL.createObjectURL(file);
    downloadLink.href = feedbackReportObjectUrl;
    downloadLink.download = `zpetna-vazba-${datePart}_${timePart}.txt`;
    downloadLink.hidden = false;
  } catch (error) {
    document.querySelector("#feedback-status").textContent =
      `Soubor .txt se nepodařilo připravit ke stažení: ${error instanceof Error ? error.message : String(error)}`;
    clearPreparedFeedbackReport();
    return;
  }
  const emailLink = document.querySelector("#feedback-email-link");
  const subject = encodeURIComponent("Zpětná vazba k aplikaci Vybavenost území");
  const body = encodeURIComponent(
    "Dobrý den,\n\nv příloze posílám export zpětné vazby k aplikaci.\n\n"
    + "Soubor .txt je potřeba přiložit ručně."
  );
  emailLink.href = `mailto:xavavrusa@vutbr.cz?subject=${subject}&body=${body}`;
  document.querySelector("#feedback-status").textContent =
    "Report je připraven. Klikněte na „Stáhnout report .txt“; poté se zpřístupní připravený e-mail.";
}

function clearNotice() {
  elements.notice.hidden = true;
  elements.notice.textContent = "";
  elements.notice.classList.remove("notice-error");
}

function updateCapacityEdit(event) {
  const input = event.target.closest("[data-record-id][data-capacity-field]");
  if (!input) return;
  const viewport = captureRequirementViewport(input);
  if (input.dataset.recordId === SPECIAL_BUILDING_PARKING_REALIZATION_KEY) {
    updateSpecialParkingCapacityEdit(event, input, viewport);
    return;
  }
  const brnoItem = BRNO_REQUIREMENTS.find((item) => item.key === input.dataset.recordId);
  if (brnoItem) {
    const field = input.dataset.capacityField;
    const numericFields = ["quantity", "override"];
    const numericValue = numericOrNull(input.value);
    if (numericFields.includes(field)
      && input.value !== "" && (numericValue === null || numericValue < 0)) {
      input.setCustomValidity("Zadejte nezápornou číselnou hodnotu nebo pole ponechte prázdné.");
      if (event.type === "change") input.reportValidity();
      return;
    }
    input.setCustomValidity("");
    if (field === "quantity") {
      const current = state.capacityEdits[brnoItem.key] || {};
      const date = input.value !== "" && !current.date ? getLocalIsoDate() : current.date;
      state.capacityEdits[brnoItem.key] = { ...current, quantity: input.value, date };
      const dateInput = input.closest(".capacity-editor")?.querySelector('[data-capacity-field="date"]');
      if (dateInput instanceof HTMLInputElement && date) dateInput.value = date;
    } else {
      const current = state.overrides[brnoItem.key] || { quantity: "", note: "", author: "", history: [] };
      state.overrides[brnoItem.key] = {
        ...current,
        [field === "override" ? "quantity" : "note"]:
          field === "override" && input.value !== "" ? Number(input.value) : input.value
      };
    }
    saveScenario();
    const card = input.closest("[data-bsp-item]");
    if (card) {
      const capacity = state.capacityEdits[brnoItem.key] || {};
      const override = state.overrides[brnoItem.key] || {};
      const result = card.querySelector("[data-brno-capacity-result]");
      if (result) result.textContent = getBrnoCapacityResultText(brnoItem, capacity, override);
      renderBrnoRequirementResult(card, brnoItem);
    }
    updateScenarioStatus();
    return;
  }
  const item = state.items.find((candidate) => candidate.key === input.dataset.recordId);
  if (!item) return;
  const id = item.key;
  const field = input.dataset.capacityField;
  const numericFields = [
    "quantity", "override", "parkingShare", "parkingLegalArea", "parkingLegalFactor",
    "accessDistance", "accessVhdMinutes"
  ];
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
    const existing = state.overrides[id] || { quantity: "", note: "", author: "", history: [] };
    state.overrides[id] = { ...existing, quantity: input.value === "" ? "" : Number(input.value) };
  } else if (field === "overrideNote") {
    const existing = state.overrides[id] || { quantity: "", note: "", author: "", history: [] };
    state.overrides[id] = { ...existing, note: input.value };
  } else if (field === "overrideAuthor") {
    const existing = state.overrides[id] || { quantity: "", note: "", author: "", history: [] };
    state.overrides[id] = { ...existing, author: input.value };
  } else {
    const existing = getCapacityEdit(item);
    state.capacityEdits[id] = { ...existing, [field]: input.value };
  }
  if (["override", "overrideNote", "overrideAuthor"].includes(field) && event.type === "change") {
    recordUserOverrideChange(item, input.closest(".requirement-card")?.dataset.kind || "");
  }
  saveScenario();
  if ((field === "parkingLegalArea" || field === "parkingLegalFactor") && event.type === "change") {
    const disclosureState = captureRequirementDisclosureState();
    renderRequirements();
    restoreRequirementDisclosureState(disclosureState);
    restoreRequirementViewport(viewport, false);
    updateScenarioStatus();
    return;
  }
  const cards = [...document.querySelectorAll(".requirement-card")]
    .filter((candidate) => candidate.dataset.itemKey === id);
  for (const card of cards) {
    const calc = adjustedResult(item, card.dataset.kind);
    const auditSummary = card.querySelector(".capacity-audit-summary");
    if (auditSummary) {
      auditSummary.textContent = getUserOverrideAuditSummary(state.overrides[id], getCapacityUnit(item));
    }
    const isDelta = card.dataset.kind === "recommended"
      && !isCompleteStandardMode() && hasStandardComparison(item)
      && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
    const isSchoolRecommendation = card.dataset.kind === "recommended"
      && SCHOOL_CAPACITY_ITEM_IDS.has(item.id);
    const comparison = isDelta ? getStandardComparison(item) : null;
    card.classList.toggle("is-zero", isDelta
      ? comparison.delta === 0
      : isHealthcareDemandItem(item)
        ? getTargetedPlanEstimate(item, card.dataset.kind === "recommended").metrics.find((metric) =>
          metric.label.startsWith(item.id === "VYB_09A_4"
            ? "Poptávka po gynekologii po odečtu ověřené kapacity"
            : "Poptávka nepokrytá ověřenou kapacitou"))?.value === 0
        : isRequirementSatisfied(calc));
    card.classList.toggle("is-capacity-satisfied",
      state.capacityEdits[id]?.satisfiesNearbyCapacity === true);
    const title = card.querySelector(".requirement-title strong");
    if (title) title.textContent = displayRequirementItemName(item, card.dataset.kind);
    const result = card.querySelector(".capacity-result");
    if (result) updateCapacityResultText(result, item, calc, card.dataset.kind);
    if (item.id === "VYB_06A") {
      const dayStatus = card.querySelector("[data-social-day-service-status]");
      if (dayStatus) dayStatus.textContent = getSocialDayServiceAccess(item).status;
    } else if (item.id === "VYB_03C") {
      const accessResult = card.querySelector('[data-spatial-access-result="sport"]');
      const targetDistance = calculateSpatialEstimate(item)?.distanceMeters;
      const actualDistance = numericOrNull(state.capacityEdits[id]?.accessDistance);
      if (accessResult) {
        accessResult.textContent = actualDistance === null || !Number.isFinite(targetDistance)
          ? "Skutečnou dostupnost zatím nelze vyhodnotit."
          : actualDistance <= targetDistance
            ? `Splňuje standard (${formatNumber(actualDistance)} m z ${formatNumber(targetDistance)} m).`
            : `Nesplňuje standard (${formatNumber(actualDistance)} m z ${formatNumber(targetDistance)} m).`;
      }
    }
    const calculation = card.querySelector(".calculation-section");
    if (calculation) {
      const replacement = renderItemCalculation(item, card.dataset.kind);
      calculation.replaceChildren(...replacement.childNodes);
    }
    const metrics = card.querySelector(".item-results");
    if (metrics) {
      metrics.replaceWith(renderItemMetrics(item, card.dataset.kind, isDelta));
    }
    if (isHealthcareDemandItem(item)) {
      const quantity = card.querySelector(".requirement-quantity");
      const projectEstimate = getTargetedPlanEstimate(item, card.dataset.kind === "recommended");
      if (quantity) quantity.textContent = formatRequirementSummaryQuantity(
        item, card.dataset.kind, getTargetedPlanSummaryText(item, projectEstimate));
      const investment = card.querySelector(".requirement-investment");
      if (investment) investment.textContent = formatCost(projectEstimate?.cost ?? null);
    }
    if (isSchoolRecommendation) {
      const quantity = card.querySelector(".requirement-quantity");
      if (quantity) quantity.textContent = formatRequirementSummaryQuantity(
        item, card.dataset.kind, getMmrSchoolSummaryText(item));
    } else if (!isHealthcareDemandItem(item)) {
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
        const source = ["override", "overrideNote", "overrideAuthor"].includes(otherInput.dataset.capacityField)
          ? state.overrides[id]
          : state.capacityEdits[id];
        const fieldKey = otherInput.dataset.capacityField === "override" ? "quantity"
          : otherInput.dataset.capacityField === "overrideNote" ? "note"
            : otherInput.dataset.capacityField === "overrideAuthor" ? "author" : otherInput.dataset.capacityField;
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
  restoreRequirementViewport(viewport);
}

function updateSpecialParkingCapacityEdit(event, input, viewport) {
  const id = SPECIAL_BUILDING_PARKING_REALIZATION_KEY;
  const field = input.dataset.capacityField;
  if (["quantity", "override"].includes(field)) {
    const value = numericOrNull(input.value);
    if (input.value !== "" && (input.validity.badInput || value === null || value < 0)) {
      input.setCustomValidity("Zadejte nezápornou číselnou hodnotu nebo pole ponechte prázdné.");
      if (event.type === "change") input.reportValidity();
      return;
    }
    input.setCustomValidity("");
  }
  if (field === "quantity") {
    const existing = state.capacityEdits[id] || {};
    const date = input.value !== "" && !existing.date ? getLocalIsoDate() : existing.date;
    state.capacityEdits[id] = { ...existing, quantity: input.value, date };
    if (date) {
      const dateInput = input.closest(".capacity-controls")?.querySelector('[data-capacity-field="date"]');
      if (dateInput instanceof HTMLInputElement) dateInput.value = date;
    }
  } else if (["source", "date", "note"].includes(field)) {
    const existing = state.capacityEdits[id] || {};
    state.capacityEdits[id] = { ...existing, [field]: input.value };
  } else {
    const existing = state.overrides[id] || { quantity: "", note: "", author: "", history: [] };
    const overrideField = field === "override" ? "quantity"
      : field === "overrideNote" ? "note" : "author";
    const value = field === "override" && input.value !== ""
      ? Number(input.value) : input.value;
    state.overrides[id] = { ...existing, [overrideField]: value };
    if (field === "override" && event.type === "change"
      && hasActiveUserOverride(state.overrides[id])) {
      const updated = state.overrides[id];
      const baseline = Number.isFinite(updated.baselineQuantity)
        ? updated.baselineQuantity : getSpecialBuildingParkingEstimate().estimatedStalls;
      const history = Array.isArray(updated.history) ? updated.history : [];
      const previousQuantity = history.length
        ? history[history.length - 1].quantity : baseline;
      history.push({
        previousQuantity,
        quantity: updated.quantity,
        author: updated.author?.trim() || "Neuvedeno",
        reason: updated.note?.trim() || "Neuveden",
        changedAt: new Date().toISOString()
      });
      state.overrides[id] = {
        ...updated,
        baselineQuantity: baseline,
        history,
        updatedAt: history[history.length - 1].changedAt
      };
    }
  }
  saveScenario();
  if (event.type !== "change") return;
  const disclosureState = captureRequirementDisclosureState();
  renderCategories();
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  restoreRequirementViewport(viewport);
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
  const sortCards = (container) => {
    const cards = [...container.querySelectorAll(":scope > .requirement-card")];
    cards.sort((left, right) =>
      Number(left.classList.contains("is-zero")) - Number(right.classList.contains("is-zero")));
    container.append(...cards);
  };
  parent.querySelectorAll(":scope > .requirement-topic-group").forEach(sortCards);
  parent.querySelectorAll(".reserve-capacity-group").forEach((group) => {
    sortCards(group);
  });
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
  BRNO_REQUIREMENTS.forEach((item) => rows.push([
    item.label, item.key, "Brněnská vyhláška", selectedKeys.has(item.key) ? "Ano" : "Ne"
  ]));
  rows.push([], ["OVĚŘENÁ KAPACITA A UŽIVATELSKÝ NÁVRH"],
    ["Vybavenost", "ID", "Ověřená kapacita", "Zdroj", "Datum", "Poznámka k ověření", "Návrhová hodnota", "Autor přepisu", "Důvod přepisu", "Výchozí hodnota modelu", "Čas poslední změny", "Historie přepisů (JSON)", "Splňuje kapacitu v okolí"]);
  state.items.forEach((item) => {
    const capacity = scenario.capacityEdits[item.key] || {};
    const override = scenario.overrides[item.key] || {};
    if (![capacity.quantity, capacity.source, capacity.date, capacity.note, capacity.satisfiesNearbyCapacity,
      override.quantity, override.author, override.note, override.updatedAt, override.history?.length]
      .some((value) => value !== undefined && value !== "" && value !== null)) return;
    rows.push([
      displayCatalogItemName(item), item.id, capacity.quantity ?? "", capacity.source ?? "",
      capacity.date ?? "", capacity.note ?? "", override.quantity ?? "", override.author ?? "", override.note ?? "",
      override.baselineQuantity ?? "", override.updatedAt ?? "",
      JSON.stringify(Array.isArray(override.history) ? override.history : []),
      capacity.satisfiesNearbyCapacity ? "Ano" : "Ne"
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
    { wch: 16 }, { wch: 36 }, { wch: 22 }, { wch: 40 }, { wch: 28 }
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

function getRealizationExportKind(item, kindOverrides = new Map()) {
  const override = kindOverrides.get(item.key);
  if (override) return override;
  if (item.id === RESIDENT_PARKING_RECOMMENDED_ID) return "review";
  if (item.id === RESIDENT_PARKING_RESERVE_ID) return "recommended";
  if (isInfrastructureReserveItem(item)) return "infrastructure";
  if (item.userAdded || REVIEW_ALTERNATIVE_IDS.has(item.id)) return "review";
  const availableKinds = ["legal", "recommended", "review"]
    .filter((kind) => getRequirementItems([item], kind).length);
  if (isHighestStandardMode() && availableKinds.includes("review")) return "review";
  if (isCompleteStandardMode() && availableKinds.includes("recommended")) {
    return "recommended";
  }
  if (availableKinds.includes("legal")) return "legal";
  return availableKinds[0] || "legal";
}

function getRealizationExportSelectionKeys() {
  const keys = new Set(getRealizationPlan().selectedKeys);
  const validKeys = new Set([
    ...state.items.filter((item) => !REMOVED_SOCIAL_SERVICE_ITEM_IDS.has(item.id))
      .map((item) => item.key),
    ...BRNO_REQUIREMENT_KEYS
  ]);
  for (const [key, selected] of Object.entries(state.realizationSelections)) {
    if (Number(selected) === 1
      && (validKeys.has(key) || key === SPECIAL_BUILDING_PARKING_REALIZATION_KEY)
      && !(BRNO_REQUIREMENT_KEYS.has(key)
        && state.capacityEdits[key]?.satisfiesNearbyCapacity === true)) {
      keys.add(key);
    }
  }
  return keys;
}

function getExportNumericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = String(value ?? "").replace(/[\s\u00a0]/g, "")
    .replace(",", ".").match(/^-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

function getRealizationExportRow(item, kind, kindOverrides = new Map()) {
  const label = item ? displayCatalogItemName(item)
    : "Parkování realizovaných specializovaných staveb (splnění vyhlášky č. 146/2024 Sb.)";
  if (!item && kind === "legal") {
    const estimate = getSpecialBuildingParkingEstimate();
    return {
      key: SPECIAL_BUILDING_PARKING_REALIZATION_KEY,
      item: null,
      kind,
      label,
      standard: "Zákonné / povinné",
      quantityValue: estimate.stalls,
      quantityText: formatValue(estimate.stalls, "stání"),
      cost: estimate.cost,
      zero: estimate.stalls === 0,
      gray: false
    };
  }
  if (!item || isInfrastructureItem(item) && !isInfrastructureReserveItem(item)) return null;
  const effectiveKind = kindOverrides.get(item.key) || kind
    || getRealizationExportKind(item, kindOverrides);
  const calc = adjustedResult(item, effectiveKind);
  const projectEstimate = getProjectPlanningEstimate(item, effectiveKind);
  const targetedEstimate = projectEstimate?.type === "targeted" ? projectEstimate : null;
  let quantityValue = getSummaryRequirementQuantity(calc);
  let quantityText = quantityValue === null || quantityValue === undefined
    ? "není vypočteno" : formatValue(quantityValue, getRequirementUnit(item));
  if (projectEstimate?.type === "infrastructure-reserve"
    || projectEstimate?.type === "infrastructure") {
    quantityValue = numericOrNull(projectEstimate.quantity);
    quantityText = formatValue(quantityValue, projectEstimate.unit || getRequirementUnit(item));
  } else if (targetedEstimate) {
    quantityText = getTargetedPlanSummaryText(item, targetedEstimate);
    quantityValue = targetedEstimate.facilityPlan?.count
      ?? getExportNumericValue(quantityText);
  } else if (projectEstimate?.type === "supermarket") {
    quantityValue = projectEstimate.quantity;
    quantityText = getTargetedPlanSummaryText(item, projectEstimate);
  } else if (projectEstimate?.type === "polyclinic") {
    quantityValue = projectEstimate.clinicCount;
    quantityText = `${formatNumber(projectEstimate.clinicCount)} ordinací`;
  } else if (projectEstimate?.type === "public-space") {
    quantityValue = projectEstimate.minimumArea;
    quantityText = formatValue(projectEstimate.minimumArea, projectEstimate.unit || "m²");
  } else if (effectiveKind === "recommended" && MMR_SCHOOL_RECOMMENDED_ITEM_IDS.has(item.id)) {
    const recommendation = getMmrSchoolRecommendation(item);
    quantityValue = recommendation?.requiredCapacity ?? quantityValue;
    quantityText = getMmrSchoolSummaryText(item);
  }
  const targetedCost = targetedEstimate?.metrics.find((metric) =>
    metric.label.toLocaleLowerCase("cs").includes("investic"))?.value
    ?? targetedEstimate?.facilityPlan?.cost;
  const cost = targetedEstimate
    ? numericOrNull(targetedCost)
    : projectEstimate?.cost ?? getDisplayedCost(item, calc, effectiveKind);
  const comparisonDelta = effectiveKind === "recommended"
    && !isCompleteStandardMode()
    && hasStandardComparison(item)
    && !SCHOOL_CAPACITY_ITEM_IDS.has(item.id)
    ? getStandardComparison(item).delta : null;
  const isDelta = comparisonDelta !== null;
  const housingEstimate = getHousingContributionEstimate(item);
  const isZero = targetedEstimate
    ? quantityValue === 0 || getExportNumericValue(quantityText) === 0
    : projectEstimate?.type === "supermarket" ? projectEstimate.quantity === 0
    : housingEstimate ? housingEstimate.municipalApartments === 0
    : projectEstimate?.type === "polyclinic" ? projectEstimate.clinicCount === 0
    : projectEstimate?.type === "infrastructure-reserve"
      || projectEstimate?.type === "infrastructure" ? projectEstimate.quantity === 0
    : projectEstimate?.type === "public-space" ? projectEstimate.minimumArea === 0
    : item.id === "VYB_10B_1" && projectEstimate?.type === "item"
      ? projectEstimate.quantity === 0
      : isDelta ? comparisonDelta === 0
      : isRequirementSatisfied(calc);
  return {
    key: item.key,
    item,
    kind: effectiveKind,
    label,
    standard: effectiveKind === "legal" ? "Zákonné / povinné"
      : effectiveKind === "recommended" ? "Metodika MMR / doporučené"
        : effectiveKind === "infrastructure" ? "Rezerva"
          : "Nadstandard / definované uživatelem",
    quantityValue,
    quantityText: formatRequirementSummaryQuantity(item, effectiveKind, quantityText),
    cost: numericOrNull(cost),
    zero: isZero,
    gray: effectiveKind === "recommended" && isMmrRequirementInactive(item)
      || state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true
  };
}

function getRealizationCompletionGroups(selectedKeys) {
  const groups = [
    { kind: "legal", title: "Zákonné / povinné", items: [] },
    { kind: "recommended", title: "Metodika MMR / doporučené", items: [] },
    { kind: "review", title: "Nadstandard / definované uživatelem", items: [] }
  ];
  const selectedParkingBase = getSelectedResidentParkingBaseItem();
  const parkingReserve = state.items.find((item) => item.id === RESIDENT_PARKING_RESERVE_ID);
  const parkingAlreadyChosen = Boolean(selectedParkingBase
    || parkingReserve && Number(state.realizationSelections[parkingReserve.key]) === 1);
  for (const group of groups) {
    const seen = new Set();
    for (const item of getRequirementItems(state.items, group.kind)) {
      if (seen.has(item.key) || selectedKeys.has(item.key)
        || HIDDEN_DUPLICATE_ITEM_IDS.has(item.id)
        || isInfrastructureItem(item)
        || isRequirementReserveItem(item)
        || item.id === "VYB_01_J"
        || state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true
        || group.kind === "recommended" && isMmrRequirementInactive(item)
        || parkingAlreadyChosen && isResidentParkingBaseItem(item)) continue;
      seen.add(item.key);
      group.items.push(item);
    }
  }
  if (!selectedKeys.has(SPECIAL_BUILDING_PARKING_REALIZATION_KEY)) {
    groups[0].items.push(null);
  }
  return groups.filter((group) => group.items.length > 0);
}

function showRealizationCompletionDialog(groups) {
  return new Promise((resolve) => {
    const dialog = createElement("dialog", "realization-export-dialog");
    const form = createElement("form", "realization-export-form");
    const heading = createElement("h2", "", "Doplnit položky před exportem?");
    form.append(heading);
    form.append(createElement("p", "realization-export-note",
      "Výběr obsahuje nepovinné položky. Zvolte, které chybějící skupiny má aplikace dopočítat a přidat do realizace. Infrastruktura se do exportu nezahrnuje."));
    const list = createElement("div", "realization-export-groups");
    for (const group of groups) {
      const fieldset = document.createElement("fieldset");
      fieldset.className = "realization-export-group";
      const legend = document.createElement("legend");
      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.exportCompletionKind = group.kind;
      const standardMode = getSelectedStandardMode();
      checkbox.checked = group.kind === "legal"
        || standardMode !== "ZAKON" && group.kind === "recommended"
        || standardMode === "MAX_STANDARD" && group.kind === "review";
      label.append(checkbox, document.createTextNode(
        `${group.title} (${group.items.length})`));
      legend.append(label);
      fieldset.append(legend);
      const itemList = createElement("ul", "realization-export-item-list");
      for (const item of group.items) {
        const row = getRealizationExportRow(item, group.kind);
        const name = item ? displayCatalogItemName(item)
          : "Parkování realizovaných specializovaných staveb (splnění vyhlášky č. 146/2024 Sb.)";
        const itemRow = createElement("li", "realization-export-item");
        itemRow.append(createElement("span", "realization-export-item-name", name));
        const detail = createElement("span", "realization-export-item-detail");
        if (row) {
          detail.append(
            createElement("span", "realization-export-item-quantity", row.quantityText),
            createElement("span", "realization-export-item-cost", formatCost(row.cost))
          );
        } else {
          detail.append(createElement("span", "", "bez vypočteného výsledku"));
        }
        itemRow.append(detail);
        itemList.append(itemRow);
      }
      fieldset.append(itemList);
      list.append(fieldset);
    }
    form.append(list);
    const actions = createElement("div", "realization-export-actions");
    const cancel = createElement("button", "text-button", "Zrušit export");
    cancel.type = "button";
    const confirmButton = createElement("button", "text-button realization-export-primary",
      "Dopočítat výběr a pokračovat");
    confirmButton.type = "submit";
    actions.append(cancel, confirmButton);
    form.append(actions);
    dialog.append(form);
    document.body.append(dialog);
    let result = null;
    const finish = (value) => {
      result = value;
      dialog.close();
    };
    cancel.addEventListener("click", () => finish(null));
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      result = [...form.querySelectorAll("[data-export-completion-kind]:checked")]
        .map((input) => input.dataset.exportCompletionKind);
      dialog.close();
    });
    dialog.addEventListener("close", () => {
      dialog.remove();
      resolve(result);
    }, { once: true });
    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      finish(null);
    });
    dialog.showModal();
  });
}

function showRealizationZeroConfirmation(rows) {
  return new Promise((resolve) => {
    const dialog = createElement("dialog", "realization-export-dialog");
    const content = createElement("div", "realization-export-form");
    content.append(createElement("h2", "", "Potvrdit nulové nebo zašedlé položky"));
    content.append(createElement("p", "realization-export-note",
      "Některé vybrané položky mají nulovou potřebu nebo jsou podle parametrů scénáře zašedlé. Chcete je přesto zahrnout do realizace a exportu?"));
    const list = createElement("ul", "realization-export-item-list realization-export-warning-list");
    rows.forEach((row) => {
      const reasons = [
        row.zero ? "vypočtená potřeba je 0" : "",
        row.gray ? "položka je zašedlá" : ""
      ].filter(Boolean).join("; ");
      list.append(createElement("li", "", `${row.label} — ${reasons}`));
    });
    content.append(list);
    const actions = createElement("div", "realization-export-actions");
    const cancel = createElement("button", "text-button", "Zrušit export");
    cancel.type = "button";
    const exclude = createElement("button", "text-button", "Vynechat tyto položky");
    exclude.type = "button";
    const include = createElement("button", "text-button realization-export-primary",
      "Ano, zahrnout a exportovat");
    include.type = "button";
    actions.append(cancel, exclude, include);
    content.append(actions);
    dialog.append(content);
    document.body.append(dialog);
    let result = null;
    const finish = (value) => {
      result = value;
      dialog.close();
    };
    cancel.addEventListener("click", () => finish(null));
    exclude.addEventListener("click", () => finish(false));
    include.addEventListener("click", () => finish(true));
    dialog.addEventListener("close", () => {
      dialog.remove();
      resolve(result);
    }, { once: true });
    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      finish(null);
    });
    dialog.showModal();
  });
}

function applyExportCompletionSelection(groups, selectedKinds) {
  const kindRank = { legal: 0, recommended: 1, review: 2 };
  const selectedCandidates = new Map();
  for (const group of groups) {
    if (!selectedKinds.includes(group.kind)) continue;
    for (const item of group.items) {
      const key = item?.key || SPECIAL_BUILDING_PARKING_REALIZATION_KEY;
      const current = selectedCandidates.get(key);
      if (!current || kindRank[group.kind] > kindRank[current.kind]) {
        selectedCandidates.set(key, { item, kind: group.kind });
      }
    }
  }
  const recommendedParking = [...selectedCandidates.values()].find((candidate) =>
    candidate.item?.id === RESIDENT_PARKING_RECOMMENDED_ID);
  const legalParking = [...selectedCandidates.values()].find((candidate) =>
    candidate.item?.id === RESIDENT_PARKING_ID
      || candidate.item?.id === RESIDENT_PARKING_LAW_ID);
  if (recommendedParking && legalParking) selectedCandidates.delete(legalParking.item.key);
  for (const candidate of selectedCandidates.values()) {
    if (!candidate.item) {
      state.realizationSelections[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] = 1;
      continue;
    }
    state.realizationSelections[candidate.item.key] = 1;
    if (isResidentParkingBaseItem(candidate.item)) {
      for (const item of state.items.filter(isResidentParkingBaseItem)) {
        if (item.key !== candidate.item.key) state.realizationSelections[item.key] = 0;
      }
      const reserve = state.items.find((item) => item.id === RESIDENT_PARKING_RESERVE_ID);
      if (reserve) state.realizationSelections[reserve.key] = 0;
    }
  }
  const kindOverrides = new Map([...selectedCandidates.values()]
    .filter((candidate) => candidate.item)
    .map((candidate) => [candidate.item.key, candidate.kind]));
  if (selectedCandidates.size) {
    renderCategories();
    renderRequirements();
    saveScenario();
    updateScenarioStatus();
  }
  return kindOverrides;
}

function getRealizationExportRows(kindOverrides = new Map()) {
  const selectedKeys = getRealizationExportSelectionKeys();
  const items = state.items.filter((item) => selectedKeys.has(item.key)
    && (!isInfrastructureItem(item) || isInfrastructureReserveItem(item)));
  const rows = [];
  const parkingItems = items.filter(isResidentParkingBaseItem);
  for (const item of items.filter((candidate) => !isResidentParkingBaseItem(candidate))) {
    const kind = getRealizationExportKind(item, kindOverrides);
    const row = getRealizationExportRow(item, kind, kindOverrides);
    if (row) rows.push(row);
  }
  const selectedParkingBase = getSelectedResidentParkingBaseItem();
  const parkingReserve = state.items.find((item) => item.id === RESIDENT_PARKING_RESERVE_ID);
  const reserveSelected = parkingReserve
    && selectedKeys.has(parkingReserve.key);
  const parkingBase = reserveSelected ? selectedParkingBase
    : parkingItems.find((item) => item.key === selectedParkingBase?.key);
  if (parkingBase) rows.push(...getResidentParkingExportRows(parkingBase));
  if (selectedKeys.has(SPECIAL_BUILDING_PARKING_REALIZATION_KEY)) {
    rows.push(getRealizationExportRow(null, "legal"));
  }
  for (const item of BRNO_REQUIREMENTS) {
    if (!selectedKeys.has(item.key)) continue;
    const estimate = getBrnoRequirementEstimate(item);
    const quantityText = !estimate || estimate.incomplete
      ? "nelze vyhodnotit"
      : item.type === "cars" ? formatValue(estimate.minimum, "stání")
        : item.type === "bikes" ? formatValue(estimate.total, "míst")
          : formatValue(estimate.spaces, "stromových prostorů");
    rows.push({
      key: item.key,
      item: null,
      kind: "brno",
      label: item.label,
      standard: "Brněnská vyhláška",
      quantityValue: estimate
        ? item.type === "cars" ? estimate.minimum
          : item.type === "bikes" ? estimate.total : estimate.spaces
        : null,
      quantityText,
      cost: numericOrNull(estimate?.cost?.cost),
      zero: estimate?.minimum === 0 || estimate?.total === 0 || estimate?.spaces === 0,
      gray: state.capacityEdits[item.key]?.satisfiesNearbyCapacity === true
    });
  }
  return rows.filter(Boolean);
}

function getParkingCostForQuantity(item, kind, quantity) {
  if (!Number.isFinite(quantity) || quantity < 0) return null;
  const buildingShare = Math.min(1, Math.max(0,
    numericOrNull(getParkingEstimateValue(item, kind, "PARKING_BUILDING_SHARE")) ?? 0));
  const volumePerStall = Math.max(0,
    numericOrNull(getParkingEstimateValue(item, kind, "PARKING_VOLUME_PER_STALL_M3")) ?? 0);
  const buildingRate = numericOrNull(getParkingEstimateValue(item, kind, "PARKING_BUILDING_RATE_M3"));
  const surfaceRate = numericOrNull(getParkingEstimateValue(item, kind, "PARKING_SURFACE_RATE_M3"));
  if (buildingRate === null || surfaceRate === null) return null;
  const surcharge = Math.max(0,
    numericOrNull(getParkingEstimateValue(item, kind, "ITEM_COST_SURCHARGE")) ?? 0.08);
  const zrn = quantity * volumePerStall
    * (buildingShare * buildingRate + (1 - buildingShare) * surfaceRate);
  return Math.round(zrn * (1 + surcharge));
}

function getResidentParkingExportRows(item) {
  const isRecommended = item.id === RESIDENT_PARKING_RECOMMENDED_ID
    || hasActiveUserOverride(state.overrides[item.key]) && item.id === RESIDENT_PARKING_ID;
  const kind = isRecommended ? "recommended" : "legal";
  const estimate = adjustedResult(item, kind);
  const totalQuantity = numericOrNull(estimate.required);
  if (!isRecommended || totalQuantity === null) {
    const row = getRealizationExportRow(item, kind);
    return row ? [row] : [];
  }
  const legalItem = state.items.find((candidate) => candidate.id === RESIDENT_PARKING_ID);
  if (!legalItem) return [];
  const legalQuantity = getResidentParkingLegalEstimate(legalItem)?.requiredStalls;
  if (!Number.isFinite(legalQuantity)) return [];
  const legalCost = getParkingCostForQuantity(legalItem, "legal", legalQuantity);
  const totalCost = getParkingCostForQuantity(item, kind, totalQuantity);
  const excessQuantity = Math.max(0, totalQuantity - legalQuantity);
  const excessCost = legalCost === null || totalCost === null
    ? null : Math.max(0, totalCost - legalCost);
  const legalRow = getRealizationExportRow(legalItem, "legal");
  const excessRow = getRealizationExportRow(item, kind);
  if (!legalRow || !excessRow) return [];
  legalRow.quantityValue = legalQuantity;
  legalRow.quantityText = formatValue(legalQuantity, "stání");
  legalRow.cost = legalCost;
  excessRow.label = `${displayCatalogItemName(item)} – část nad zákonné minimum`;
  excessRow.standard = isRecommended
    ? "Doporučené / vlastní navýšení"
    : excessRow.standard;
  excessRow.quantityValue = excessQuantity;
  excessRow.quantityText = formatValue(excessQuantity, "stání");
  excessRow.cost = excessCost;
  excessRow.zero = excessQuantity === 0;
  return [legalRow, excessRow];
}

function createRealizationExportSheet(rows, omittedUnpricedCount) {
  const output = [
    ["PŘEHLED VYBAVENOSTI K REALIZACI"],
    ["Položka", "Standard", "Kapacita / rozsah", "Vypočtená cena (Kč)"],
    ...rows.map((row) => [row.label, row.standard, row.quantityText, row.cost])
  ];
  const total = rows.reduce((sum, row) => sum + row.cost, 0);
  output.push([], ["CELKEM", "", "", total]);
  if (omittedUnpricedCount) {
    output.push([], [`Vynecháno položek bez vypočtené ceny: ${omittedUnpricedCount}`]);
  }
  const sheet = XLSX.utils.aoa_to_sheet(output);
  sheet["!cols"] = [{ wch: 68 }, { wch: 34 }, { wch: 32 }, { wch: 24 }];
  sheet["!autofilter"] = { ref: `A2:D${rows.length + 2}` };
  for (let rowIndex = 2; rowIndex < rows.length + 2; rowIndex += 1) {
    const cell = sheet[`D${rowIndex + 1}`];
    if (cell) cell.z = '#,##0 "Kč"';
  }
  sheet[`D${rows.length + 4}`].z = '#,##0 "Kč"';
  return sheet;
}

async function exportRealizedItemsWorkbook() {
  if (!state.workbook || !state.items.length) {
    showNotice("Nejprve načtěte zdrojový Excel, aby bylo možné vyexportovat realizované položky.", true);
    return;
  }
  const selectedKeys = getRealizationExportSelectionKeys();
  if (!selectedKeys.size) {
    showNotice("Nejprve zaškrtněte v přehledu alespoň jednu položku k realizaci.", true);
    return;
  }
  const selectedItems = state.items.filter((item) => selectedKeys.has(item.key)
    && !isInfrastructureItem(item));
  const hasSpecialParking = selectedKeys.has(SPECIAL_BUILDING_PARKING_REALIZATION_KEY);
  const hasBrnoRequirements = BRNO_REQUIREMENTS.some((item) => selectedKeys.has(item.key));
  const hasReserve = state.items.some((item) =>
    selectedKeys.has(item.key) && isRequirementReserveItem(item));
  if (!selectedItems.length && !hasSpecialParking && !hasBrnoRequirements && !hasReserve) {
    showNotice("Vybranou technickou infrastrukturu nelze exportovat; vyberte její rezervu nebo jinou položku vybavenosti.", true);
    return;
  }
  const selectedNonLegal = selectedItems.some((item) =>
    getRealizationExportKind(item) !== "legal");
  let kindOverrides = new Map();
  if (selectedNonLegal) {
    const groups = getRealizationCompletionGroups(selectedKeys);
    if (groups.length) {
      const selectedKinds = await showRealizationCompletionDialog(groups);
      if (selectedKinds === null) return;
      kindOverrides = applyExportCompletionSelection(groups, selectedKinds);
    }
  }
  let rows = getRealizationExportRows(kindOverrides);
  const questionableRows = rows.filter((row) => row.zero || row.gray);
  if (questionableRows.length) {
    const confirmed = await showRealizationZeroConfirmation(questionableRows);
    if (confirmed === null) return;
    if (!confirmed) {
      const omittedKeys = new Set(questionableRows.map((row) => row.key));
      rows = rows.filter((row) => !omittedKeys.has(row.key));
    }
  }
  const pricedRows = rows.filter((row) => Number.isFinite(row.cost));
  const omittedUnpricedCount = rows.length - pricedRows.length;
  if (!pricedRows.length) {
    showNotice(hasBrnoRequirements
      ? "Brněnské předpisy stanovují kapacity, nikoli ceny. Bez samostatného ocenění nelze tyto položky zahrnout do cenového exportu."
      : "Pro vybrané položky není dostupná žádná vypočtená cena; export nebyl vytvořen.", true);
    return;
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook,
    createRealizationExportSheet(pricedRows, omittedUnpricedCount), "Realizované položky");
  const file = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
  const download = document.createElement("a");
  const url = URL.createObjectURL(new Blob([file], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  }));
  download.href = url;
  download.download = `${state.fileName.replace(/\.(?:xlsx|xls|xlsm)$/i, "") || "Projekt"}_realizace.xlsx`;
  document.body.append(download);
  download.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    download.remove();
  }, 1000);
  const omittedNote = omittedUnpricedCount
    ? ` Vynecháno položek bez vypočtené ceny: ${omittedUnpricedCount}.` : "";
  showNotice(`Export realizací obsahuje ${pricedRows.length} oceněných položek; technická infrastruktura byla vynechána.${omittedNote}`);
}

function toggleNearbyCapacitySatisfaction(event) {
  const button = event.target.closest("button[data-capacity-satisfaction-item-key]");
  if (!button) return false;
  event.preventDefault();
  event.stopPropagation();
  const viewport = captureRequirementViewport(button);
  if (button.dataset.capacitySatisfactionItemKey === SPECIAL_BUILDING_PARKING_REALIZATION_KEY) {
    const disclosureState = captureRequirementDisclosureState();
    const existing = state.capacityEdits[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] || {};
    state.capacityEdits[SPECIAL_BUILDING_PARKING_REALIZATION_KEY] = {
      ...existing,
      satisfiesNearbyCapacity: !existing.satisfiesNearbyCapacity
    };
    saveScenario();
    renderCategories();
    renderRequirements();
    restoreRequirementDisclosureState(disclosureState);
    restoreRequirementViewport(viewport);
    updateScenarioStatus();
    return true;
  }
  if (BRNO_REQUIREMENT_KEYS.has(button.dataset.capacitySatisfactionItemKey)) {
    const itemKey = button.dataset.capacitySatisfactionItemKey;
    const existing = state.capacityEdits[itemKey] || {};
    state.capacityEdits[itemKey] = {
      ...existing,
      satisfiesNearbyCapacity: !existing.satisfiesNearbyCapacity
    };
    const disclosureState = captureRequirementDisclosureState();
    saveScenario();
    renderCategories();
    renderRequirements();
    restoreRequirementDisclosureState(disclosureState);
    restoreRequirementViewport(viewport);
    updateScenarioStatus();
    return true;
  }
  const item = state.items.find((candidate) =>
    candidate.key === button.dataset.capacitySatisfactionItemKey);
  if (!item) return true;
  const disclosureState = captureRequirementDisclosureState();
  const existing = getCapacityEdit(item);
  state.capacityEdits[item.key] = {
    ...existing,
    satisfiesNearbyCapacity: !existing.satisfiesNearbyCapacity
  };
  saveScenario();
  renderCategories();
  renderRequirements();
  restoreRequirementDisclosureState(disclosureState);
  restoreRequirementViewport(viewport);
  updateScenarioStatus();
  return true;
}

document.querySelector("#excel-export").addEventListener("click", () => {
  try {
    exportProjectWorkbook();
  } catch (error) {
    showNotice(`Export projektu do Excelu se nepodařil: ${error instanceof Error ? error.message : String(error)}`, true);
  }
});

elements.realizedExport.addEventListener("click", async () => {
  try {
    await exportRealizedItemsWorkbook();
  } catch (error) {
    showNotice(`Export realizovaných položek se nepodařil: ${error instanceof Error ? error.message : String(error)}`, true);
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
    loaded.items.push(...createInfrastructureReserveVariants(loaded.items));
    Object.assign(state, loaded, {
      selectedCategory: "", selectedItemKey: "", search: "", hideUnrealizedItems: false,
      collapsedCategories: new Set(), collapsedLists: {}, projectInputGroupsOpen: {},
      derivedDetailsOpen: false, capacityEdits: {}, overrides: {},
      reviewChecks: {}, socialServiceChecks: {}, standardComparisons: {},
      healthChecks: {}, parkingShares: {},
      realizationSelections: {},
      itemEstimateValues: {},
      brnoInputs: {},
      scenarioChanged: false, inputsChanged: false, persistenceError: false
    });
    const scenarioReadError = readScenario(loaded.projectScenario);
    elements.search.value = state.search;
    elements.excelExport.disabled = false;
    elements.realizedExport.disabled = false;
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
  state.socialServiceChecks = {};
  state.standardComparisons = {};
  state.healthChecks = {};
  state.parkingShares = {};
  state.realizationSelections = {};
  state.itemEstimateValues = {};
  state.brnoInputs = {};
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
elements.derivedGrid.addEventListener("change", handleInputChange);
elements.legalList.addEventListener("change", handleInlineProjectInputChange);
elements.recommendedList.addEventListener("change", handleInlineProjectInputChange);
elements.infrastructureList.addEventListener("change", handleInlineProjectInputChange);
elements.reviewList.addEventListener("change", handleInlineProjectInputChange);
elements.categoryList.addEventListener("click", (event) => {
  if (event.target.closest("[data-realization-toggle-all]")) {
    toggleAllRealizations();
    return;
  }
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
  const feedbackButton = event.target.closest("button[data-feedback-context]");
  if (feedbackButton) {
    event.preventDefault();
    event.stopPropagation();
    openFeedbackReport(feedbackButton.dataset.feedbackContext);
    return;
  }
  if (toggleNearbyCapacitySatisfaction(event)) return;
  const realizationScopeButton = event.target.closest(
    "button[data-realization-group-toggle], button[data-realization-section-toggle]"
  );
  if (realizationScopeButton) {
    const scope = realizationScopeButton.hasAttribute("data-realization-group-toggle")
      ? realizationScopeButton.closest(".requirement-topic-group")
      : realizationScopeButton.closest(".requirement-column");
    if (scope) toggleRealizationScope(scope);
    return;
  }
  const realizationButton = event.target.closest("button[data-realization-item-key]");
  if (realizationButton) {
    event.preventDefault();
    event.stopPropagation();
    if (BRNO_REQUIREMENT_KEYS.has(realizationButton.dataset.realizationItemKey)) {
      setBrnoRequirementSelection(
        realizationButton.dataset.realizationItemKey,
        realizationButton.getAttribute("aria-pressed") !== "true"
      );
      return;
    }
    if (realizationButton.dataset.realizationItemKey
      === SPECIAL_BUILDING_PARKING_REALIZATION_KEY) {
      setSpecialParkingRealizationSelection(
        realizationButton.getAttribute("aria-pressed") !== "true");
      return;
    }
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
  renderBrnoRequirementList();
  saveScenario(false);
});
document.querySelector(".derived-details").addEventListener("toggle", (event) => {
  state.derivedDetailsOpen = event.currentTarget.open;
  saveScenario(false);
});
document.querySelector("#open-user-guide").addEventListener("click", startUserGuide);
document.querySelector("#guide-back").addEventListener("click", () => {
  showUserGuideStep(userGuideStepIndex - 1);
});
document.querySelector("#guide-next").addEventListener("click", () => {
  if (userGuideStepIndex === USER_GUIDE_STEPS.length - 1) finishUserGuide();
  else showUserGuideStep(userGuideStepIndex + 1);
});
document.querySelector("#guide-exit").addEventListener("click", finishUserGuide);
document.addEventListener("keydown", handleUserGuideKeydown);
elements.requirementColumns.addEventListener("input", updateTargetedPlanInput);
elements.requirementColumns.addEventListener("change", updateSpecialBuildingParkingInput);
elements.brnoList.addEventListener("input", updateBrnoRequirementInputs);
elements.brnoList.addEventListener("change", updateBrnoRequirementInputs);
elements.brnoList.addEventListener("input", updateCapacityEdit);
elements.brnoList.addEventListener("change", updateCapacityEdit);
elements.categoryList.addEventListener("change", updateRealizationSelection);
elements.legalList.addEventListener("input", updateCapacityEdit);
elements.legalList.addEventListener("change", updateCapacityEdit);
elements.legalList.addEventListener("input", updateHealthcareCheck);
elements.legalList.addEventListener("change", updateHealthcareCheck);
elements.legalList.addEventListener("input", updateManualReview);
elements.legalList.addEventListener("change", updateManualReview);
elements.legalList.addEventListener("input", updateSocialServiceAvailability);
elements.legalList.addEventListener("change", updateSocialServiceAvailability);
elements.legalList.addEventListener("input", updateStandardComparison);
elements.legalList.addEventListener("change", updateStandardComparison);
elements.legalList.addEventListener("click", handleRequirementClick);
elements.recommendedList.addEventListener("input", updateCapacityEdit);
elements.recommendedList.addEventListener("change", updateCapacityEdit);
elements.recommendedList.addEventListener("input", updateSocialServiceAvailability);
elements.recommendedList.addEventListener("change", updateSocialServiceAvailability);
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
elements.infrastructureList.addEventListener("click", handleRequirementClick);
elements.reserveList.addEventListener("input", updateCapacityEdit);
elements.reserveList.addEventListener("change", updateCapacityEdit);
elements.reserveList.addEventListener("input", updateManualReview);
elements.reserveList.addEventListener("change", updateManualReview);
elements.reserveList.addEventListener("change", handleInlineProjectInputChange);
elements.reserveList.addEventListener("click", handleRequirementClick);
elements.reviewList.addEventListener("input", updateCapacityEdit);
elements.reviewList.addEventListener("change", updateCapacityEdit);
elements.reviewList.addEventListener("input", updateManualReview);
elements.reviewList.addEventListener("change", updateManualReview);
elements.reviewList.addEventListener("input", updateHealthcareCheck);
elements.reviewList.addEventListener("change", updateHealthcareCheck);
elements.reviewList.addEventListener("input", updateSocialServiceAvailability);
elements.reviewList.addEventListener("change", updateSocialServiceAvailability);
elements.reviewList.addEventListener("click", handleRequirementClick);
elements.resetInputs.addEventListener("click", resetScenario);
document.querySelector("#open-feedback-report").addEventListener("click", () => {
  openFeedbackReport();
});
document.querySelector("#feedback-form").addEventListener("submit", exportFeedbackReport);
document.querySelector("#feedback-download-link").addEventListener("click", () => {
  document.querySelector("#feedback-email-link").hidden = false;
  document.querySelector("#feedback-status").textContent =
    "Stažení reportu bylo zahájeno. V případě potřeby otevřete připravený e-mail, přiložte .txt a zprávu odešlete.";
});
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
