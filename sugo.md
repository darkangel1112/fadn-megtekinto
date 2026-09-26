# Súgó

## Mire való ez az app?

A FADN Megtekintő az ömlesztett Excel-exportot a táblázatos sablon alapján olvasható, visszaépített nézetté alakítja.

## Használat

1. Tölts fel egy ömlesztett Excel-fájlt.
2. Tölts fel egy táblázatos sablon Excel-fájlt.
3. Válassz üzemet és munkalapot.
4. Ha kell, kapcsold be a szűréseket vagy a technikai oszlopok megjelenítését.
5. Az aktuális nézet CSV-be le is tölthető.

Az elfogadott forrásfájlok kiterjesztése: `.xlsx`, `.xlsm` vagy `.xls`. A két fájlnak ugyanahhoz az adatállapothoz és sablonverzióhoz kell tartoznia.

## Célzott export

A bal oldali sávban a feltöltött ömlesztett és sablonfájl után megjelenik a Célzott export gomb.

- Válaszd ki az exportálandó üzemeket.
- Az aktív exporttárgyak alapból ki vannak jelölve; amelyik nem kell, annak a jelölését vedd ki.
- Az első elérhető export az ősszel vetett terület a 5C táblából.
- Az export tartalmazza a mezei leltár összesítő és releváns vetési sorait, köztük az m5430 „Tavaszi vetések előkészítése” sort is.
- Csak a nem üres és nem nulla záró értékű sorok kerülnek bele.
- A búza külön sorai megmaradnak, és ha legalább az egyik értékes, egy összesített Búza sor is készül.
- Egy üzem esetén egy Excel-fájl, több üzem esetén üzemenként külön Excel-fájl ZIP-csomagban tölthető le.
- Az exportált értékek numerikusak, két tizedessel jelennek meg.
- A célzott export gomb a dialog megnyitása és az export elkészítése közben is látható marad a bal oldali sávban.
- A célzott export ablaka az elkészítés és a letöltés után is nyitva marad; a Bezárás gomb, a jobb felső `×` és az Esc billentyű is szabályosan bezárja.
- A „Földterületi adatok” export a `t1_a` munkalap nem nulla és nem üres adatsort tartalmazó sorait adja vissza.
- A földterületi munkalap a megjelenítő sorrendjét követi: `FarmCode`, `RowCode`, `RowTitle`, majd a hét t1_a értékoszlopot.
- A t1_a összesítő és számított sorai is bekerülnek, ha van bennük adat.
- A „Készletek” export az 5C és 6B zárókészleteit külön `Készletek` munkalapon kapcsolja össze.
- A készlet-párosítás az `mXXXX` kódszabály alapján történik, az összesítő sorok nélkül.
- A `6B - 5C` különbözet vásárolt készletként jelenik meg, a csak 6B-ben szereplő készletsorok is bekerülnek.
- Az „Állatok” export a 6A munkalap nem nulla záróállományait (`osz = 12`) külön `Állatok` munkalapra teszi.
- Az állat-export oszlopai: `Üzemkód`, `RowCode`, `Sor megnevezése`, `Mértékegység`, `Záróérték`.
- A 6A összesítő sorai kimaradnak, a nem nulla `előző sor t-ban` súlysorok viszont megmaradnak.
- A ló-, hal- és egyéb állatsorok is bekerülnek, ha a sablonban szerepelnek és van nem nulla záróértékük.
- A normál állatlétszám `db`, a súlysor `t`, a méhcsaládok `család` mértékegységet kapnak.
- Ha egy kiválasztott exporttárgyhoz nincs nem nulla adat, a munkalap csak fejléccel készül; üres Excel-tábla és üres szűrő nem kerül rá.
- A célzott export dialog szélesebb változatban jelenik meg, hogy az üzemlista és az exporttárgyak kényelmesebben elférjenek.
- Az exporttárgyak melletti kérdőjel rövid magyarázatot ad az adott export tartalmáról.
- A súgóablak szélesebb, saját görgetősávval rendelkezik, a fejléc pedig görgetés közben is látható marad.

## Fontos

- A feltöltött fájlok nem kerülnek tartós alkalmazás-adattárba.
- Az adatok az aktuális munkamenetben élnek; a források törlése vagy a munkamenet megszűnése után nem maradnak elérhetők.
