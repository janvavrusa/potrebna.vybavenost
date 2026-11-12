# Návod k používání aplikace Vybavenost území

## Interaktivní průvodce

V horní liště klikněte na **Spustit návod**. Jednotlivá okénka vás provedou načtením sešitu, vyplněním a kontrolou vstupů, výběrem vybavenosti, úpravou výpočtu, ověřením okolní kapacity, exportem a nahlášením chyby. Můžete během čtení pracovat i se zvýrazněnou částí aplikace. Tlačítky **Zpět** a **Další** postupujete ručně; poslední krok lze dokončit tlačítkem **Dokončit** nebo kdykoli zavřít **Ukončit návod**. Fungují také šipky vlevo/vpravo a klávesa Escape. Zvýrazněná oblast se při postupu posune do zobrazení a potřebné výpočetní části se dočasně rozbalí; po ukončení se vrátí jejich původní stav rozbalení.

## 1. Otevření a načtení dat

Otevřete `index.html` v aktuálním prohlížeči. Aplikace se pokusí načíst výchozí pracovní sešit umístěný vedle stránky. Pokud sešit nenačte nebo chcete použít jiný, klikněte na **Import** a vyberte soubor `.xlsx`, `.xls` nebo `.xlsm`; soubor lze také přetáhnout na oblast Importu.

Po načtení zkontrolujte název sešitu a stav v horní liště. Počty položek a pravidel se zobrazí také jako zpráva aplikace. Při chybě načtení vyberte sešit ručně.

## 2. Zadání projektu

V horní části upravte projektové vstupy. Jsou rozdělené do skupin, které lze rozbalit a sbalit. Prázdné nebo volitelné hodnoty nevyplňujte odhadem, pokud je nemáte k dispozici. V části **Odvozené ukazatele** najdete automaticky dopočtené hodnoty a případné volitelné přepisy.

Současný počet obyvatel řešeného území, přírůstek projektu a obyvatelé spádových sídel mají v modelech odlišné role; aplikace je sčítá pouze tam, kde to příslušný výpočet vyžaduje. Obec či městskou část a typ území vyplňte podle skutečného místa záměru. Po změně vstupů se výsledky přepočítají.

## 3. Hledání a filtrování vybavenosti

- V levém panelu procházejte tematické skupiny; šipkou skupinu rozbalíte nebo sbalíte.
- Do vyhledávacího pole napište název, kód nebo zdroj.
- Kliknutím na název skupiny omezíte výstup na její položky. Kliknutím znovu filtr zrušíte.
- Kliknutím na název konkrétní položky zobrazíte její související výstupy; dalším kliknutím se vrátíte k širšímu výběru.
- Přepínač **Skrýt nerealizované** ponechá ve výsledcích pouze položky označené k realizaci.

Výstup je rozdělen na zákonné/povinné položky, doporučení MMR, technickou infrastrukturu, rezervy, další kontroly a samostatné propočty Brněnské vyhlášky. Zařazení ve výstupu samo o sobě neznamená, že je nutné položku postavit.

## 4. Výběr realizací a čtení výsledků

Barevné tlačítko u karty nebo políčko v katalogu označuje, zda má být položka zahrnuta do záměru. Výběr je uživatelský scénář; nerozhoduje o právní povinnosti. Výběr hromadně pro skupinu nebo část výstupu používejte jen po kontrole, co obsahuje.

Kliknutím na kartu ji rozbalte. V souhrnu uvidíte modelovanou potřebu a případné orientační náklady. Uvnitř karty:

1. přečtěte stručný popis a výsledky;
2. otevřete **Výpočet / jak jsme k výsledku došli**;
3. zkontrolujte použité vstupy, mezikroky, jednotky a zdroje;
4. podle potřeby upravte dostupné předpoklady.

U škol, polikliniky a některých dalších položek jsou dostupné zvláštní výměrové nebo cenové odhady. Karty **Brněnská vyhláška** obsahují samostatné propočty parkování motorových vozidel, jízdních kol a prostor pro růst stromů. Tyto odhady vycházejí z editovatelných vstupů a nenahrazují ověření konkrétní stavby, účinného znění předpisu ani rozpočtu.

## 5. Ověření okolní kapacity

V kartě otevřete část **Ověřený stav v okolí a uživatelský návrh** (název se může u konkrétního typu lišit):

- automatické hledání okolních zařízení je pouze pomůcka; nalezený záznam ověřte u poskytovatele nebo z důvěryhodného zdroje;
- zadejte volnou kapacitu ve správné jednotce a podle možnosti doplňte zdroj a datum;
- prázdné pole znamená neověřeno, nikoli nulovou kapacitu;
- funkce pro potvrzení dostatečné kapacity je uživatelský záznam kontroly, ne automatické zjištění aplikace;
- uživatelský návrh a jeho důvod lze zadat samostatně.

U zdravotnických položek a dostupnosti může být potřeba ověřit vzdálenost nebo čas dojezdu podle skutečné trasy. Internetové výsledky samy o sobě nepotvrzují pěší dostupnost ani provozní kapacitu.

## 6. Ukládání a export

Změny se průběžně ukládají do místního scénáře tohoto prohlížeče; nejsou automaticky odesílány na server. Pro přenos práce použijte **Export Excel**. Exportovaný sešit uchovejte jako pracovní kopii a při dalším použití jej znovu načtěte.

Tlačítko **Export realizací (Excel)** připraví zvláštní přehled vybraných realizací. Před exportem zkontrolujte seznam, varianty parkování, nulové nebo nedoložené hodnoty a případná upozornění.

Tlačítko **Obnovit hodnoty ze sešitu** odstraní místní změny scénáře a vrátí vstupy i předpoklady k hodnotám výchozího sešitu. Před použitím exportujte práci, kterou chcete zachovat.

## 7. Nahlášení chyby nebo návrhu

U každé položky je tlačítko **Nahlásit chybu u této položky**. Kliknutí otevře formulář a předvyplní místo v aplikaci, které lze upravit. Nahlásit lze také obecný problém přes tlačítko v horní liště.

Do formuláře napište:

- co jste očekávali a co se stalo;
- kroky, kterými lze problém zopakovat;
- případně správnou hodnotu, zdroj nebo návrh změny.

Po odeslání formuláře klikněte na **Stáhnout report .txt**. Až se zahájí stahování, zpřístupní se **Otevřít připravený e-mail**; e-mailový klient se otevře až tímto samostatným krokem, aby se stahování reportu nepřerušilo. Soubor přiložte ručně a e-mail odešlete sami; aplikace nic neposílá automaticky. Pokud se e-mailový klient neotevře, vytvořený soubor přiložte k nové zprávě ručně. Zkontrolujte obsah reportu před odesláním, protože může obsahovat zadané poznámky, hodnoty a název sešitu.

## 8. Interpretace a omezení

Modelové potřeby, vzdálenosti, kapacity, výměry a ceny jsou pracovní výpočty a mohou záviset na vstupních předpokladech. U sociálních služeb jsou demografické koeficienty návrhové scénáře, nikoli statistické údaje nebo normativy. U Brněnské vyhlášky ověřte její účinné znění, přechodná ustanovení, územní/regulační plán i konkrétní projekt. Odhadované sazby neurčují cenu stavby.

Před rozhodnutím o projektu ověřte použitelná pravidla a podklady u příslušných institucí, správců sítí, poskytovatelů a projektantů.
