# Postup pro přidání nebo změnu položky vybavenosti

Tento postup je určen pro změny položek v aplikaci i v pracovním katalogu. Záznam musí být dohledatelný od zdroje přes předpoklady a výpočet až po výsledek, který uživatel uvidí a může podle pravidel editovat. Pracovní sešit `Seznam Vybavenosti_pracovni_katalog_pravidel.xlsx` a aplikační logika `app.js` jsou společně zdrojem dat a chování; změna pouze v jednom z nich není dokončená, pokud výsledek závisí na obou.

## 1. Určete, zda je nová položka skutečně potřeba

- Projděte existující položky, jejich varianty, duplicitní záznamy, pravidla a výpočty v pracovním sešitu i aplikaci. Použijte existující položku nebo její variantu, pokud to významově odpovídá požadavku; nevytvářejte druhou kartu jen kvůli jinému popisku.
- Zkontrolujte všechny standardy: **Zákonem / Povinné**, **Metodika MMR / Doporučené**, **Nadstandard / Definované uživatelem** a případné **Rezervy**. Rozhodněte, do kterého výstupu každá varianta patří a zda je položka alternativou, doplňkem, nebo samostatnou povinností.
- Stanovte stabilní kód `VYB_*`, skupinu, název, jednotku a rozsah. Před přidělením kódu zkontrolujte, že nekoliduje s existujícími položkami; pokud se stejné kódy v sešitu používají ve více tematických skupinách, zachovejte jejich jednoznačné tematické rozlišení.

## 2. Ověřte pravidlo a zdroje

U každého právního nebo metodického tvrzení uveďte přesný zdroj, verzi či datum účinnosti, konkrétní místo a přímý odkaz nebo název přiloženého souboru.

Upřednostňujte zdroje v tomto pořadí:

1. Platné znění právních předpisů v [oficiální e-Sbírce](https://www.e-sbirka.cz) a příslušné prováděcí předpisy.
2. Oficiální metodiky a standardy příslušného ministerstva ([MMR](https://mmr.gov.cz)); pro pravidla dostupnosti zejména místní soubor [TB050MMR01 – Standardy dostupnosti veřejné infrastruktury](./TB050MMR01-Standardy-dostupnosti-verejne-infrastruktury-aktualizace-2020-03.pdf).
3. Oficiální statistiky a oborová data státních institucí ([ČSÚ](https://www.czso.cz), [ÚZIS](https://www.uzis.cz), [MPSV](https://www.mpsv.cz)) pro populační a kapacitní předpoklady.
4. Ověřené místní podklady (pasporty, kapacity, územní plán, provozovatelé sítí) pro stav konkrétního území.
5. Cenové podklady RTS/JKSO nebo jiný dohledatelný cenový zdroj pro investiční odhad; používejte konkrétní odpovídající skupinu a konstrukci, např. z místního souboru [RTS Speciální stavby](./RTS%20Speci%C3%A1ln%C3%AD%20stavby.pdf).

- Citujte ustanovení, tabulku/řádek či stranu, podmínky použití a datum ověření. Rozlišujte povinnost, doporučení, empirický údaj a modelový předpoklad.
- Neoznačujte doporučení nebo odhad za zákonnou povinnost. Pokud důvěryhodný zdroj chybí, uveďte to výslovně a označte vstup jako editovatelný scénář; nedoplňujte domnělou normu ani falešně přesnou cenu.
- Zaznamenejte omezení: např. typ území, populační práh, spádovou oblast, dochozí/dojezdovou vzdálenost, minimální hustotu, provozní režim či skutečnost, že místní kapacitu musí ověřit uživatel.

## 3. Založte konzistentní záznam v pracovním sešitu

Upravujte existující listy a jejich záhlaví; nevytvářejte paralelní „pomocný“ list bez potřeby a neměňte názvy záhlaví používané aplikací.

- **`Model_Vybavenosti`**: položku vložte do správné tematické skupiny za řádek `CLUSTER`. Vyplňte stabilní ID, název, kapacitní poptávku, kapacitní a prostorové počty, jednotku, návrhový počet, plochy a cenu jen tam, kde mají doložitelný význam. Zachovejte jednotky a návaznost vzorců. Nevyplněná nebo neznámá hodnota není nula.
- **`Pravidla_a_vzorce`**: každé použitelné zákonné, doporučené nebo uživatelské pravidlo uveďte jako samostatný záznam se správným ID vybavenosti, typem, podmínkou, přesným vzorcem/algoritmem, vstupy a jednotkami, výstupem, zaokrouhlením, zdrojem, konkrétním místem, stavem ověření a omezením.
- **`Slovnik_vstupu`**: přidejte každý skutečně potřebný vstup s unikátním ID/kódem, názvem, rolí, typem, jednotkou, výchozí hodnotou, validací, vzorcem/odvozením, místem použití, informací o editovatelnosti a zdrojem. U uživatelsky měnitelného předpokladu vysvětlete, co jeho změna ovlivňuje.
- Aktualizujte související vzorce v `Vstupy_a_KPI` či `Konstanty_a_Ciselniky` jen tehdy, když je položka skutečně potřebuje. Neodkazujte na neexistující buňky, listy, ID pravidel ani vstupy.
- Ověřte vzorce, rozsahy a formát jednotek přímo v Excelu/kompatibilním tabulkovém procesoru; nenechávejte výsledek závislý na neuloženém nebo nepřepočítaném vzorci.

## 4. Propojte data s aplikační logikou

`readWorkbook()` načítá položky z `Model_Vybavenosti` a pravidla a vstupy z jejich uvedených listů. Zachovejte požadované záhlaví definované v `MODEL_HEADERS`, `RULE_HEADERS` a `INPUT_HEADERS`.

- Ověřte, že se položka načte právě jednou, spadá do správné tematické skupiny, má správnou jednotku a její pravidla se přiřadí správné položce i standardu. Při nejednoznačných nebo opakovaných ID doplňte a ověřte mapování; nepředpokládejte, že samotné ID vždy určí položku.
- Rozlišujte položku převzatou ze zdroje od položky vložené uživatelem bez zákonného/metodického podkladu. Pro uživatelskou položku nastavte všechny příslušné klasifikační značky a texty; její přítomnost sama o sobě nedokládá povinnost.
- Výpočtovou logiku držte v datové/výpočtové vrstvě (zejména `getProjectPlanningEstimate()` a souvisejících výpočtových helper funkcích), ne pouze v renderovací funkci. Každý výsledek musí mít pojmenovaný typ, jasné jednotky a definované chování pro prázdný vstup, nulu a neplatnou hodnotu.
- Pokud položka potřebuje zvláštní model, doplňte i všechny související množiny a mapování v `app.js` (např. klasifikaci infrastruktury, uživatelských položek, pravidel, rezerv a duplicit). Zvláštní/virtuální položku nevkládejte současně jako běžnou načtenou položku i samostatnou kartu.
- U nové technické infrastruktury prověřte také bilance a rezervy. U nové běžné vybavenosti prověřte realizaci a export; vynechání infrastruktury z exportu běžných položek nesmí omylem vynechat její výslovně zvolené rezervy.

## 5. Zobrazte úplný a srozumitelný výpočet

- Hned pod názvem po otevření položky uveďte krátké pravidlo: čím se tato varianta liší od zákonné/doporučené alternativy a jaký hlavní práh nebo způsob dostupnosti se na ni vztahuje. Náhled karty ponechte stručný.
- V části **Hlavní výsledky návrhového výpočtu** zobrazte jen výsledky, které skutečně vyplývají z doložených dat nebo jasně označeného modelu: kapacitu/počet, relevantní plochu či výkon a vypočtenou cenu, má-li podklad.
- V části **Výpočet / jak jsme k výsledku došli** zobrazte čitelně vzorec, mezikroky, jednotky, zaokrouhlení, původ každého vstupu a omezení. Hodnoty navržené jako scénář označte jako modelové a udělejte je editovatelné podle stávajících UI vzorů. Parametr, který není možné věrohodně určit, neschovávejte za automatický odhad.
- Výstup jasně odlišuje návrhovou potřebu od uživatelem ověřené místní kapacity. Není-li dostupnost prokázána automaticky, ponechte kontrolu uživateli a nepředstírejte, že odhad počtu bodů potvrzuje skutečnou pěší trasu nebo provozní kapacitu.
- Zachovejte jednotné zobrazení editovatelných vstupů, zdrojů, validačních chyb, šedých/neaktivních položek a stavů realizace v detailu, přehledu, datovém katalogu a exportu.

## 6. Ověřte celou cestu uživatele

Před předáním ověřte alespoň následující:

1. Spusťte `node --check .\app.js` a `git diff --check`.
2. Načtěte skutečný pracovní sešit v aplikaci. Ověřte správný název, tematickou skupinu, standard, jednotku, pravidla a zdrojové popisy.
3. Porovnejte alespoň jeden ruční výpočet s aplikací. Otestujte prahovou hodnotu z obou stran, nulu, prázdný vstup, editaci vstupu a přepnutí typu území/standardu, pokud na nich pravidlo závisí.
4. Ověřte změnu výsledku po editaci a uložení/obnově scénáře. Zkontrolujte, že výsledky v kartě, přehledu a exportu jsou konzistentní a položka se nezobrazuje duplicitně.
5. Otestujte zařazení do realizace, nulový nebo neaktivní stav, související infrastrukturu či rezervu a export, pokud jsou pro položku relevantní.
6. Zkontrolujte světlý a tmavý režim a úzké zobrazení. Poté zkontrolujte diff a potvrďte, že změna nezahrnula nesouvisející soubory.

Dokončení znamená, že je doložen zdroj i klasifikace, výpočet je přezkoumatelný, editovatelné předpoklady jsou správně zapojené a všechny relevantní výstupní plochy zobrazují stejnou logiku.
