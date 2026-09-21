# Változásnapló

## 2026-09-21 – Célzott export gomb állapotkezelése

- a Célzott export gomb a dialog megnyitása és az export elkészítése közben is látható marad;
- a jobb felső `×`-szel vagy Esc billentyűvel történő bezárás is törli a dialog állapotát;
- a gomb eltűnéséből adódó félrevezető főnézeti állapot megszűnik.

## 2026-09-21 – Üres célzott exportlapok kezelése

- az üres mezei leltár-, készlet- és állat-exportlapok fejléc-only munkalapként készülnek;
- üres munkalapokra nem kerül Excel-táblaobjektum vagy AutoFilter;
- az üres `Állatok` lap nem hoz létre hibás `tbl_Allatok` táblát.

## 2026-09-21 – Munkalapválasztó hover-kezelés

- a hosszú munkalapcímek rövidített, stabil szélességű címkét kapnak a legördülő menüben;
- a teljes munkalapcím továbbra is megjelenik a fő nézetben;
- a javítás megszünteti a hosszú 6A/6B címek miatti menü-újraméretezést és túlcsordulást.

## 2026-09-21 – 6A állatok export

- aktiválva az „Állatok” exporttárgy;
- új `Állatok` munkalap a 6A `osz = 12` záróállományaiból;
- az összesítő sorok kizárása a kettős elszámolás elkerülésére;
- a nem nulla `előző sor t-ban` súlysorok megtartása;
- ló-, hal- és egyéb állatsorok automatikus kezelése a sablon alapján;
- mértékegység jelölése (`db`, `t`, `család`);
- valódi Excel-tábla létrehozása `tbl_Allatok` néven;
- a célzott export dialog szélesebb, nagy méretű változata.

## 2026-09-21 – Dokumentációs állapot

- hozzáadva a projekt README-je;
- hozzáadva az architektúra- és adatfolyam-leírás;
- hozzáadva az export-specifikáció;
- hozzáadva az üzemeltetési útmutató;
- hozzáadva az érzékeny adatok nélküli tesztelési útmutató;
- rögzítve a jelenlegi 5C, 5C/6B készlet- és 6A állat-export állapota.

## 2026-09-21 – 5C/6B készletek export

- aktiválva a „Készletek” exporttárgy;
- új `Készletek` munkalap az üzemenkénti Excel-kimenetben;
- szándékos `mXXXX` kódcsalád-szabály szerinti 5C–6B párosítás;
- összesítő sorok kizárása;
- saját és vásárolt készlet számítása;
- csak 6B-ben szereplő részletes sorok kezelése;
- `6B < 5C` ellenőrzés piros hibaállapottal és nullázott számított értékekkel;
- egy munkafüzetben több kiválasztott export-munkalap támogatása.

## 2026-09-21 – Célzott export és súgó UI finomítás

- az aktív exporttárgyak alapértelmezés szerint kijelölve jelennek meg;
- magyarázó kérdőjel került az őszi vetett terület exporttárgya mellé;
- a célzott export dialog megnyitása a fő táblázat kirajzolása után történik;
- a súgó dialog szélesebb, belső görgethető tartalomterületet kapott.

## 2026-09-20 – Célzott 5C export

- célzott export dialog;
- üzemek kijelölése és tömeges kijelölés kezelése;
- 5C mezei leltár nem nulla záróértékeinek exportja;
- `m5410`, `m5411`–`m5418`, `m5422`–`m5425`, `m5429` és `m5430` mapping;
- külön m5411/m5412 sorok és összesített `Búza` sor;
- együzemes Excel és többüzemes ZIP kimenet;
- numerikus Excel-értékek, két tizedes, szűrő, rácsvonal és vékony keret;
- exportkészítési spinner és letöltés a nyitva maradó dialogban;
- fő nézeti állapot megőrzése a célzott export dialog megnyitása és bezárása után.

## Alapverzió

- ömlesztett Excel és táblázatos sablon betöltése;
- sablon szerinti munkalapnézet;
- üzem- és munkalapválasztás;
- kitöltött, záró/összesítő és technikai oszlop szűrők;
- aktuális nézet CSV-letöltése;
- beépített magyar súgó.
