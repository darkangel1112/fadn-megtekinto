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
- [ ] A fő üzem-, munkalap- és szűrőállapot megmarad a dialog megnyitásakor.
- [ ] A dialog bezárása után a fő tábla változatlan marad.
- [ ] Az export létrehozása alatt spinner látható.
- [ ] A letöltés az elkészült export után megjelenik.
- [ ] A dialog a letöltés után nyitva marad.
- [ ] Az aktív exporttárgyak alapból kijelölve jelennek meg.
- [ ] Az „Ősszel vetett terület” és a „Készletek” kérdőjele rövid magyarázatot jelenít meg.
- [ ] Az „Állatok” exporttárgy alapból kijelölve jelenik meg, ha a 6A munkalap elérhető.
- [ ] Az „Állatok” kérdőjele röviden jelzi a 6A záróállomány, az összesítő-kizárás és a súlysorok szabályát.
- [ ] A célzott export dialog az előző változathoz képest szélesebb, nagy méretű változatban jelenik meg.
- [ ] A súgó dialog szélesebb, a tartalma saját görgetősávval görgethető, a „Súgó” fejléc látható marad.
- [ ] A célzott export dialog bezárása után az első munkalapváltás azonnal frissíti a fő táblát.

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

### Üres célzott exportlapok

- [ ] Üres mezei leltár esetén a munkalap fejléc-only marad, hibajelzés nélkül.
- [ ] Üres készlet esetén a munkalap fejléc-only marad, hibajelzés nélkül.
- [ ] Üres állatadat esetén nincs üres `tbl_Allatok` táblázatdefiníció.

## Automatizált tesztállapot

A repóban jelenleg nincs külön `tests/` könyvtár vagy CI-folyamat. A fejlesztés során statikus ellenőrzés, Streamlit AppTest és valódi forrásfájlokból végzett, repón kívüli ellenőrzés használható, de érzékeny adatot ezekből sem szabad a repóba menteni.

Új automatizált tesztek hozzáadásakor szintetikus sablon- és ömlesztett adatot kell létrehozni, amely csak a szükséges sor- és oszlopstruktúrát tartalmazza.
