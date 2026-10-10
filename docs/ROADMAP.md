# Fejlesztési terv

## Elhalasztott: ömlesztett Excel betöltés gyorsítása

**Státusz:** elnapolva, későbbi fejlesztési lehetőség.

A jelenlegi ömlesztett Excel-beolvasás saját XML-parserrel működik. Nagy fájloknál a fő teljesítményigényt a `_parse_sheet_rows` Python-szintű XML-, cella- és oszlopfeldolgozása okozza.

Későbbi fejlesztésként vizsgálandó a `python-calamine` elsődleges Excel-olvasóként történő bevezetése. A megvalósításnál a jelenlegi parser átmeneti tartalékként maradjon meg, amíg a tényleges mintafájlokkal és szintetikus nagy fájlokkal nem igazolt az eredményazonosság.

Visszatéréskor szükséges feladatok:

- a `python-calamine` kompatibilitásának ellenőrzése a helyi és szerveroldali Python-környezetben;
- `Path` és memóriabeli `bytes` forrásból történő betöltés tesztelése;
- fejléc-, üres cella-, numerikus és szöveges értékek paritásvizsgálata;
- az összes meglévő nézet- és exporteredmény összehasonlítása;
- benchmark valós és 50 000–100 000 soros szintetikus fájlokon;
- a régi XML-parser csak sikeres migrációs tesztek után legyen kivezethető.

A `lxml`-es parsercsere és a regex-optimalizáció másodlagos lehetőségként kezelendő, nem elsődleges megoldásként.
