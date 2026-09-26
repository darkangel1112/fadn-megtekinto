# Tesztelési útmutató

## Adatvédelmi alapelv

Partner- vagy üzemazonosítókat tartalmazó forrásfájl nem lehet a repóban. A jelenlegi projektben ezért nincs valódi Excel-tesztfixture és nincs olyan tesztfájl, amely érzékeny adatot tárolna.

Teszteléshez:

- szintetikus, anonimizált munkafüzetet használj;
- a fixture-t a repón kívül tárold;
- vagy memóriából adj át tesztadatot;
- a teszt után az ideiglenes fájlt töröld.

## Gyors statikus ellenőrzés

A projekt gyökeréből:

```powershell
python -m py_compile app.py src\workbook_mapper.py
git diff --check
```

Az első parancs a Python-szintaxist, a második a whitespace- és patch-hibákat ellenőrzi.

## Alkalmazásindítási ellenőrzés

```powershell
python -m streamlit run app.py --server.headless true --server.address 127.0.0.1 --server.port 8503
(Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8503/_stcore/health).Content
```

Az egészségügyi végpont elvárt válasza: `ok`.

## Kézi elfogadási ellenőrzőlista

### Források és fő nézet

- [ ] Forrásfájlok nélkül a célzott export gomb letiltott.
- [ ] Mindkét fájl feltöltése után a célzott export gomb aktív.
- [ ] Üzemváltáskor a fő tábla az adott üzem adatait mutatja.
- [ ] Munkalapváltáskor a sablon szerinti nézet jelenik meg.
- [ ] Hosszú munkalapcím fölé húzva az egérmutatót a legördülő menü nem ugrik el és az opció kiválasztható marad.
- [ ] A fő nézet szűrői működnek.

### Dialog és állapotmegőrzés

- [ ] A dialog csak a célzott export gombra kattintva jelenik meg.
- [ ] A Célzott export gomb a dialog megnyitása, az export elkészítése és a letöltés alatt is látható marad.
- [ ] A fő üzem-, munkalap- és szűrőállapot megmarad a dialog megnyitásakor.
- [ ] A dialog bezárása után a fő tábla változatlan marad.
- [ ] Az export létrehozása alatt spinner látható.
- [ ] A letöltés az elkészült export után megjelenik.
- [ ] A dialog a letöltés után nyitva marad.
- [ ] Az aktív exporttárgyak alapból kijelölve jelennek meg.
- [ ] Az „Ősszel vetett terület” és a „Készletek” kérdőjele rövid magyarázatot jelenít meg.
- [ ] Az „Állatok” exporttárgy alapból kijelölve jelenik meg, ha a 6A munkalap elérhető.
- [ ] Az „Állatok” kérdőjele röviden jelzi a 6A záróállomány, az összesítő-kizárás és a súlysorok szabályát.
- [ ] A „Földterületi adatok” exporttárgy alapból kijelölve jelenik meg, ha a t1_a munkalap elérhető.
- [ ] A „Földterületi adatok” kérdőjele jelzi a nem nulla és nem üres t1_a sorok szabályát.
- [ ] A célzott export dialog az előző változathoz képest szélesebb, nagy méretű változatban jelenik meg.
- [ ] A súgó dialog szélesebb, a tartalma saját görgetősávval görgethető, a „Súgó” fejléc látható marad.
- [ ] A célzott export dialog bezárása után az első munkalapváltás azonnal frissíti a fő táblát.
- [ ] A dialog bezárása a Bezárás gombbal, a jobb felső `×`-szel és az Esc billentyűvel is visszaállítja a fő nézetet.

### 5C export

- [ ] `m5410` és a releváns vetési sorok megjelennek, ha nem nulla értékük van.
- [ ] `m5430` „Tavaszi vetések előkészítése” néven megjelenik, ha van értéke.
- [ ] Nulla és üres értékű sorok nem kerülnek az exportba.
- [ ] `m5411` és `m5412` külön sor marad.
- [ ] Legalább az egyik búzasor esetén létrejön az összesített `Búza` sor.
- [ ] Egy üzem esetén `.xlsx`, több üzem esetén `.zip` készül.
- [ ] Az `Érték` cellák numerikusak és `0.00` formátumúak.
- [ ] A fejlécszűrő, rácsvonal és vékony cellakeret megmarad.

### Készletek export

- [ ] A „Készletek” exporttárgy aktív, ha az 5C és 6B munkalap is betöltődött.
- [ ] A készletek külön `Készletek` munkalapra kerülnek.
- [ ] Az 5C és 6B sorok az `mXXXX` kódszabály alapján párosodnak.
- [ ] Az összesítő sorok nem jelennek meg.
- [ ] A nem nulla 5C vagy 6B záróértékű részletes sorok megjelennek.
- [ ] A `m6510 – Vásárolt takarmánykeverék` csak 6B-ben szereplő sorai üres 5C mezőkkel jelennek meg, ha van értékük.
- [ ] A saját és vásárolt készlet számítása helyes.
- [ ] A `6B < 5C` eset piros hibajelzést és nullázott számított értékeket ad.
- [ ] Több kiválasztott exporttárgy esetén ugyanabba az üzemenkénti Excel-fájlba kerülnek a munkalapok.
- [ ] A készletértékek valódi numerikus cellák és két tizedes formátumúak.

### Állatok export

- [ ] Az „Állatok” külön `Állatok` munkalapra kerül.
- [ ] Csak a 6A `osz = 12` záróállományából származó, nem nulla sorok jelennek meg.
- [ ] Az összesítő sorok, köztük a részösszesítő `m6243`, nem jelennek meg.
- [ ] A nem nulla `előző sor t-ban` sorok megmaradnak `t` mértékegységgel.
- [ ] A normál állatlétszám `db`, a méhcsaládok `család` mértékegységgel jelennek meg.
- [ ] A ló- és halsorok bekerülnek, ha a sablonban szerepelnek és az értékük nem nulla.
- [ ] A `Záróérték` valódi numerikus cella, két tizedes formátummal.
- [ ] A fejlécszűrő, rácsvonal, rögzített fejléc és vékony cellakeret megmarad.
- [ ] Üres 6A adat esetén az `Állatok` lap csak fejléccel készül, Excel-táblaobjektum és AutoFilter nélkül.

### Földterületi adatok export

- [ ] A `t1_a` munkalapból külön `Földterületi adatok` munkalap készül.
- [ ] Csak azok a t1_a sorok jelennek meg, amelyekben legalább egy értékoszlop nem üres és nem nulla.
- [ ] A t1_a összesítő és számított sorai megmaradnak, ha van bennük adat.
- [ ] A hét értékoszlop a megjelenítő sorrendjében és valódi numerikus Excel-cellaként jelenik meg.
- [ ] A `FarmCode`, `RowCode` és `RowTitle` mezők megmaradnak; a technikai oszlopok és az üres sablonoszlop kimaradnak.
- [ ] A t1_a export a teljes kijelölt export negyedik munkalapja.
- [ ] Üres t1_a adat esetén a munkalap csak fejléccel készül, Excel-táblaobjektum és AutoFilter nélkül.

### Támogatási jogcímek export

- [ ] A támogatási export csak akkor választható, ha a `t7_c` és `t7_b1` forrássablon is betöltődött.
- [ ] Minden kiválasztott üzemnél megjelenik a rögzített teljes jogcímlista, üres vagy nulla érték esetén is.
- [ ] Az AKG-, erdészeti és Natura 2000 összegek a `t7_c` `osz = 3` értékeit használják.
- [ ] A CIS-YF `m7407` összege a `t7_b1` `osz = 4` eFt mezőjéből jön, nem az `osz = 3` egységadatból.
- [ ] Az eFt értékek numerikus Ft összeggé válnak (`× 1 000`) egész Ft megjelenítéssel.
- [ ] Az AKG összesítő sorok külön `Összesítő` jelölést kapnak, és nem számítódnak hozzá a részjogcímekhez.
- [ ] A forrásbeli nulla, a kitöltetlen érték, a hiányzó forrássor és az azonosítatlan kód eltérő állapotként jelenik meg.
- [ ] A kistermelői jogcím kód nélkül is szerepel, de az összege üres marad és ezt az adatállapot jelzi.
- [ ] Az `m7213` és `m7240` induló támogatási sorok nem olvadnak be az `m7407` CIS-YF sorba.
- [ ] Ha mind az öt exporttárgy be van jelölve, a `Támogatási jogcímek` az ötödik lap.
- [ ] A támogatási exporttárgy önmagában is helyes nevű `.xlsx` fájlt ad.

### Üres célzott exportlapok

- [ ] Üres mezei leltár esetén a munkalap fejléc-only marad, hibajelzés nélkül.
- [ ] Üres készlet esetén a munkalap fejléc-only marad, hibajelzés nélkül.
- [ ] Üres állatadat esetén nincs üres `tbl_Allatok` táblázatdefiníció.

## Automatizált tesztállapot

A repóban nincs CI-folyamat. A `tests/` könyvtár automatizált, szintetikus adatokkal futó célzott export-regressziós teszteket tartalmaz. Futtatás:

```powershell
python -m unittest discover -s tests -v
```

Érzékeny forrásadat vagy partnerazonosító ne kerüljön a repóba.
