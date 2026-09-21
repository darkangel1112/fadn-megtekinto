# FADN Megtekintő

A FADN Megtekintő az ömlesztett FADN Excel-exportot a táblázatos sablon alapján olvasható, visszaépített nézetté alakítja.

Az alkalmazás Streamlit-alapú, és a feltöltött munkafüzetekből dolgozik. A célzott export a 5C mezei leltárát, az 5C és 6B alapján képzett készleteket, valamint a 6A állatállomány nem nulla záróértékeit exportálja Excelbe.

## Jelenlegi állapot

### Működő funkciók

- ömlesztett Excel és táblázatos sablon feltöltése;
- üzem és munkalap kiválasztása;
- rekonstruált munkalapnézet a sablon sor- és oszlopneveivel;
- „Csak kitöltött sorok” szűrés;
- kísérleti záró-/összesítő sor szűrés;
- technikai oszlopok megjelenítése;
- aktuális nézet CSV-letöltése;
- célzott 5C export Excelbe;
- célzott készlet-export az 5C és 6B munkalap alapján;
- célzott állat-export a 6A munkalap záróállományából;
- egy üzem esetén egy Excel-fájl, több üzem esetén ZIP-csomag üzemenként külön Excel-fájlokkal;
- Excel-szűrő, rácsvonal, vékony cellakeret, fejlécformázás és numerikus két tizedes formátum.

## Gyors indítás Windows alatt

A projekt gyökérkönyvtárából:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m streamlit run app.py --server.headless true --server.address 127.0.0.1 --server.port 8503
```

Ezután az alkalmazás itt érhető el:

```text
http://127.0.0.1:8503
```

## Használat

1. Töltsd fel az ömlesztett Excel-fájlt.
2. Töltsd fel a táblázatos sablon Excel-fájlt.
3. Válassz üzemet és munkalapot.
4. Szükség esetén kapcsold be a nézeti szűrőket.
5. A célzott exporthoz kattints a bal oldali „Célzott export” gombra.
6. Jelöld ki az üzemeket és az aktív exporttárgyat.
7. Készítsd el, majd töltsd le az exportot.

A célzott export ablaka az elkészítés után nyitva marad, és csak a „Bezárás” gombbal zárható be.

## Bemeneti fájlok

Az alkalmazás a következő kiterjesztéseket fogadja el:

- `.xlsx`
- `.xlsm`
- `.xls`

Az ömlesztett fájl első munkalapjának a következő oszlopokat kell tartalmaznia:

| Oszlop | Jelentés |
| --- | --- |
| `akod` | üzemkód |
| `sor` | sorazonosító, például `m5410` |
| `osz` | oszlop-/értéktípus kódja |
| `ertek` | érték |
| `dimenzio1` | dimenzió vagy bontási érték |
| `dyn_row_serial` | dinamikus sorazonosító |

A táblázatos sablonban a kód a munkalapokat, a sorazonosítókat, a sorneveket és az oszlopkódokat olvassa ki. A részletes adat- és mapping-leírás a [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) fájlban található.

## Célzott export röviden

Az első célzott export a 5C munkalapból az ősszel vetett/mezei leltárhoz tartozó, nem üres és nem nulla záró mennyiségi értékeket veszi ki.

- Kimeneti munkalap: `Ősszel vetett terület`
- Oszlopok: `Üzemkód`, `RowCode`, `Sor megnevezése`, `Érték`
- Az értékek Excelben numerikusak, formátumuk `0.00`.
- A búza külön sorai megmaradnak, és szükség esetén keletkezik egy külön `Búza` összesítő sor.
- Az `m5430` („Tavaszi vetések előkészítése”) sor is része a jelenlegi mappingnek.

Teljes specifikáció: [docs/EXPORTS.md](docs/EXPORTS.md).

## Készletek export röviden

A `Készletek` munkalap az 5C saját zárókészletét és a 6B összes zárókészletét kapcsolja össze. A párosítás elsődleges kulcsa a sablonban kialakított `mXXXX` szabály:

- `m55xx` → `m64xx`;
- `m56xx` → `m65xx`;
- `m57xx` → `m66xx`;
- `m58xx` → `m67xx`.

Az export nem tartalmaz összesítő sorokat, csak olyan részletes sorokat, ahol az 5C vagy a 6B záróérték nem nulla. A `6B - 5C` különbözet vásárolt készletként jelenik meg. A csak 6B-ben szereplő `m6510 – Vásárolt takarmánykeverék` is bekerül, üres 5C mezőkkel.

## Állatok export röviden

A célzott export `Állatok` munkalapja a 6A sablon nem nulla záróállományait (`osz = 12`) tartalmazza.

- Oszlopok: `Üzemkód`, `RowCode`, `Sor megnevezése`, `Mértékegység`, `Záróérték`.
- A munkalap valódi Excel-táblájának technikai neve `tbl_Allatok`.
- Az összesítő sorok kimaradnak, hogy ne történjen kettős elszámolás.
- A nem nulla `előző sor t-ban` technikai sorok megmaradnak, `t` mértékegységgel.
- A normál állatlétszám `db`, a méhcsaládok `család` mértékegységgel jelennek meg.
- A ló-, hal- és egyéb állatsorok nincsenek külön fajlistával korlátozva: a sablon minden nem összesítő 6A sora feldolgozható.
- A `Záróérték` valódi numerikus Excel-cellaként, két tizedessel kerül kiírásra.
- Ha egy üzemnél nincs releváns nem nulla állatadat, az `Állatok` munkalap csak fejléccel készül, hibás üres Excel-tábla nélkül.

## Adatvédelem és tesztadatok

A feltöltött munkafüzetek nem kerülnek tartós alkalmazás-adattárba. Az adatok az aktuális Streamlit-munkamenetben vannak jelen; a források törlése vagy a munkamenet megszűnése után nem maradnak elérhetők.

Éles vagy partnerazonosítókat tartalmazó Excel/CSV-fájl nem kerülhet a Git-repóba. Tesztekhez szintetikus, anonimizált mintafájlokat kell használni a repón kívül.

## Dokumentációs térkép

- [Súgó](sugo.md) – rövid felhasználói leírás az alkalmazáson belül.
- [Architektúra](docs/ARCHITECTURE.md) – modulok, adatfolyam és bővítési pontok.
- [Export-specifikáció](docs/EXPORTS.md) – jelenlegi 5C, 5C/6B készlet- és 6A állat-export.
- [Üzemeltetés](docs/OPERATIONS.md) – helyi és systemd-alapú indítás, naplózás, hibakeresés.
- [Tesztelés](docs/TESTING.md) – érzékeny adatok nélküli ellenőrzési eljárás.
- [Változásnapló](CHANGELOG.md) – funkcionális változások és fejlesztési állapot.

## Licenc és belső használat

A repóhoz jelenleg nem tartozik külön licencfájl. A felhasználási és terjesztési szabályokat a projekt gazdája határozza meg.
