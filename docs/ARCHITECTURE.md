# Architektúra és adatfolyam

## Áttekintés

Az alkalmazás két Excel-forrásból dolgozik:

1. az ömlesztett adatfájlból;
2. a táblázatos sablonból.

A sablon adja a munkalapok, sorok és oszlopok emberi jelentését. Az ömlesztett adatok a sor- és értékkódok alapján kerülnek a sablon szerinti nézetbe.

```text
feltöltött Excel-fájlok
        │
        ├── táblázatos sablon → munkalap-, sor- és oszlopmapping
        │
        └── ömlesztett adat → normalizált rekordok
                                  │
                                  └── sor kód alapján munkalaphoz kapcsolás
                                             │
                                             ├── általános táblázatos nézet
                                             └── célzott exportnézet
```

## Fő modulok

### `app.py`

Felelős:

- Streamlit-oldal és oldalsáv kirajzolása;
- fájlfeltöltés és munkamenet-kezelés;
- fő nézeti szűrők;
- súgó- és célzott export dialog;
- CSV-, többmunkalapos Excel- és ZIP-kimenet;
- célzott export állapotának megőrzése.

### `src/workbook_mapper.py`

Felelős:

- Excel-források olvasása;
- sablonmunkalapok és sorok feltérképezése;
- ömlesztett rekordok normalizálása;
- sorazonosítók munkalaphoz kapcsolása;
- általános nézet építése;
- 5C mezei leltár célzott exportnézetének építése;
- 5C–6B kódalapú készlet-párosítás és célzott készlet-exportnézet építése;
- 6A záróállomány-alapú, összesítő sorokat kizáró állat-exportnézet építése.
- t1_a földterületi exportnézet építése a megjelenítő sorrendje és értékszűrése szerint;
- a rögzített támogatási jogcímlista előállítása a `t7_c` és `t7_b1` forráskódjai alapján, hiányzó értékek jelölésével.

### `sugo.md`

Az alkalmazáson belüli súgó forrásfájlja. Az `app.py` a projekt gyökeréből olvassa be.

## Excel-beolvasás

Az ömlesztett fájlt a mapper XML-szinten olvassa. A jelenlegi megvalósítás az ömlesztett fájl első munkalapját dolgozza fel, a shared string értékeket és a cellahivatkozások oszloppozícióit is kezeli.

A sablon betöltése `openpyxl` read-only, data-only módban történik. A sablonban:

- az első sorból származik a munkalap címe;
- a harmadik sorból az oszlopkódok;
- a negyedik sorból az oszlopcímkék;
- az ötödik sortól a C oszlop a sorazonosító, a D a sor megnevezése, az E a dinamikus sorazonosító.

## Normalizált adatmodell

Az ömlesztett adatból a mapper ezeket a mezőket tartja meg:

```text
akod             üzemkód
sor              sorazonosító
osz              értékoszlop-/értéktípus kódja
ertek            nyers érték szöveges formában
dimenzio1        első dimenzió
dyn_row_serial   dinamikus sorazonosító
sheet_name       a sablonból feloldott munkalapnév
```

Az `attach_sheet_names` a sablon `row_to_sheet` mappingje alapján tölti ki a `sheet_name` mezőt. Az ismeretlen sorazonosítók üres munkalapnévvel maradnak, ezek száma az oldalsávban figyelmeztetésként jelenik meg.

## Általános nézet

A `build_sheet_view` egy üzemhez és egy sablonmunkalaphoz:

1. kiválasztja a megfelelő `akod` és `sheet_name` rekordokat;
2. kulcsot képez a `sor`, `dyn_row_serial`, `dimenzio1` és `osz` mezőkből;
3. a sablon sorrendjében létrehozza a sorokat;
4. kitölti a sablon szerinti értékoszlopokat;
5. alkalmazza a kitöltött- és zárósor-szűrést.

A dinamikus sorazonosítók numerikusan rendeződnek, ha numerikus értéket tartalmaznak.

## t1_a földterületi export

A `build_land_view` a `t1_a` sablonhoz tartozó általános nézetből indul ki, és üzemenként csak azokat a sorokat tartja meg, amelyekben legalább egy valódi értékoszlop nem üres és nem nulla. A sablon összesítő és számított sorai nem kapnak külön kizárási szabályt.

A kimenet a `FarmCode`, `RowCode`, `RowTitle` mezőket és a nem üres kódú t1_a értékoszlopokat tartalmazza. A `DynRowSerial`, `Dimension1` és a sablon üres záró fejléce nem kerül a célzott munkalapra. Az értékek export előtt numerikus értékké alakulnak, a hiányzó értékek pedig üresek maradnak.

## Munkamenet- és dialogállapot

A betöltött forrásokat a Streamlit `session_state` tárolja. A forrásfájlok SHA-1-alapú aláírása alapján csak forrásváltozáskor történik új beolvasás.

A fő nézet vezérlői stabil kulcsokat használnak:

- `main_farm_code`;
- `main_selected_sheet`;
- `main_filled_only`;
- `main_closing_only`;
- `main_show_hidden_technical`.

Ez fontos, mert a célzott export dialog fragmentként működik, és bezáráskor újrarenderelés történik. A célzott export gomb a dialog mögött is kirajzolva marad, a `on_dismiss` callback pedig a gomb- és exportállapotot a Bezárás, a jobb felső `×` és az Esc bezárásakor is törli. A célzott export megnyitását a fő nézet és annak táblázata után kell végrehajtani, hogy a dialog bezárása után a munkalapválasztás és a fő tábla ugyanabban a teljes alkalmazás-újrarenderelésben frissüljön.

## Bővítési pontok

Új célzott export hozzáadásakor:

1. külön mapper-függvény készüljön a `workbook_mapper.py` fájlban;
2. a mapping és a forrásérték-kódok legyenek név szerint dokumentálva;
3. az export dataframe oszlopai legyenek explicit módon megadva;
4. az Excel-formázó újrahasznosítható maradjon;
5. az UI-ban az exporttárgy csak akkor legyen aktív, ha a mögöttes mapper és validáció elkészült;
6. valódi partnerfájl ne kerüljön tesztfixture-ként a repóba.

## Készlet-párosítás

A `build_stock_view` nem soronként karbantartott párosítási táblát használ. Az 5C sorazonosítóból a kódcsalád szabálya alapján állítja elő a várt 6B sorazonosítót (`m55 → m64`, `m56 → m65`, `m57 → m66`, `m58 → m67`).

A sablon összesítő sorait a megnevezés alapján kiszűri, majd a záróértékeket üzem, sorazonosító és `osz` szerint aggregálja. A számított saját és vásárolt készletet az export dataframe állítja elő. A csak 6B-ben megtalált részletes sorok külön 6B-only rekordként kerülnek az exportba.

Az exportáló réteg egy üzemenkénti munkafüzetbe több kiválasztott munkalapot ír. Több üzem esetén ezeket a munkafüzeteket ZIP-csomagba helyezi.

## Támogatási jogcímek export

A `build_support_view` a támogatási űrlap rögzített sorlistáját állítja elő üzemenként, ezért az üres vagy nulla forrásérték sem tüntet el jogcímet. Az AKG- és erdészeti összegek a `t7_c` `osz = 3` mezőjéből, a fiatal gazdák CIS-YF összege a `t7_b1` `osz = 4` mezőjéből származik. A forrás eFt értékeit a mapper Ft-ra szorozza át. Az AKG összesítő sorok a részjogcímektől megkülönböztetve szerepelnek; összeadásuk a részjogcímekkel kettős elszámolást okozna.

A célűrlapon szereplő kistermelői támogatáshoz nincs azonosított FADN-sorkód. A sor megmarad, de az összege üres és az adatállapot jelzi a hiányzó kódot. A fiatal gazdák induló támogatásának külön `m7213` és `m7240` kódjai nem kerülnek automatikusan a CIS-YF jogcímhez.

## Állat-export

A `build_animal_view` a `t6_a` sablon munkalap sorrendjében dolgozik. A 6A mozgásoszlopai közül kizárólag az `osz = 12` érték jelenti a záróállományt, ezért csak ezt az értéket olvassa.

Az export:

- csak a nem nulla záróértékű sorokat tartja meg;
- az `összesen` és `mindösszesen` megnevezésű összesítő sorokat kihagyja;
- a nem nulla `előző sor t-ban` sorokat megtartja, mert ezek az előző állatcsoport záró súlyadatai;
- nem fajlistával dolgozik, hanem a sablon minden nem összesítő 6A sorát képes feldolgozni, így a ló- és halsorok is automatikusan működnek;
- a normál állatlétszámot `db`, a súlysorokat `t`, a `m6323` méhcsalád-sort `család` mértékegységgel jelöli.

Az export dataframe oszlopai: `Üzemkód`, `RowCode`, `Sor megnevezése`, `Mértékegység`, `Záróérték`. A `RowCode` közvetlenül a 6A `mXXXX` azonosítója, ezért nincs szükség közelítő szöveges párosításra.

## Ismert korlátok

- Az ömlesztett fájlból jelenleg az első munkalap kerül feldolgozásra.
- Az ismeretlen sorazonosítók nem kerülnek automatikusan új munkalaphoz.
- A projektben jelenleg nincs automatizált, repóban tárolt tesztfixture.
