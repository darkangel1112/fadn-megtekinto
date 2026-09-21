# Export-specifikáció

## Közös szabályok

A célzott export a bal oldali sávból nyitható meg, de csak akkor használható, ha az ömlesztett adat és a táblázatos sablon is betöltődött.

A dialogban:

- az összes betöltött üzemkód kijelölhető;
- van összes kijelölése és kijelölés törlése művelet;
- a jelenleg aktív exporttárgyak alapból kijelölve jelennek meg;
- az exporttárgyak melletti kérdőjel rövid magyarázatot ad az adott munkalapról;
- több üzem kijelölése esetén üzemenként külön Excel készül;
- több Excel esetén ZIP-csomag készül;
- az export elkészítése közben spinner jelenik meg;
- az elkészült kimenet ugyanabban a dialogban tölthető le;
- a dialog a letöltés után is nyitva marad.
- ha egy kiválasztott exporttárgyhoz az adott üzemben nincs nem nulla adat, a munkalap fejlécével megmarad, de üres Excel-tábla és üres AutoFilter nem kerül rá;
- a súgó nagyobb, belső görgethető tartalomterületet használ, így a súgó fejléce látható marad.

## Aktív export: 5C mezei leltár

### Forrás

- Forrásmunkalap: `t5_c`;
- értéktípus: `osz = 5`, amely a jelenlegi forrásfájlokban a záróérték mennyiségi oszlopa;
- csak nem üres és nem nulla érték kerül exportba;
- az értékek két tizedesre kerekítve, numerikus Excel-cellaként kerülnek a kimenetbe.

### Kimeneti munkalap

```text
Ősszel vetett terület
```

### Kimeneti oszlopok

| Oszlop | Tartalom |
| --- | --- |
| `Üzemkód` | a kiválasztott üzem kódja |
| `RowCode` | a forrás sorazonosítója, például `m5411` |
| `Sor megnevezése` | exportban használt magyar megnevezés |
| `Érték` | numerikus záróérték, `0.00` Excel-formátumban |

### Aktív sorazonosítók

| RowCode | Exportmegnevezés |
| --- | --- |
| `m5410` | Mezei leltár összesen |
| `m5411` | Őszi búza |
| `m5412` | Durumbúza |
| `m5413` | Rozs |
| `m5414` | Őszi árpa |
| `m5415` | Triticale |
| `m5416` | Repce |
| `m5417` | Őszi takarmánykeverék |
| `m5418` | Évelő pillangósok |
| `m5422` | Egyéb szántóföldi kultúra |
| `m5423` | Rét-legelő |
| `m5424` | Zöldségtermelés |
| `m5425` | Virág- és dísznövény termelés |
| `m5429` | Egyéb kertészeti termelés |
| `m5430` | Tavaszi vetések előkészítése |

### Búza-kezelés

Az `m5411` és `m5412` külön sorokként megmaradnak. Ha legalább az egyik sor értéket tartalmaz, a mapper egy további sort is létrehoz:

```text
Üzemkód    RowCode    Sor megnevezése    Érték
14/1040              Búza               m5411 + m5412
```

Ennek az összesítő sornak nincs `RowCode` értéke, mert nem közvetlenül a sablon egyetlen sorát képviseli.

### Excel-formázás

Az exportált munkafüzet:

- fejlécszűrőt tartalmaz;
- az első sort rögzíti;
- a rácsvonalakat bekapcsolva hagyja;
- vékony cellakeretet használ;
- formázott fejlécet és munkalapfület kap;
- az `Érték` oszlopot `0.00` formátumban írja.

Az Excel a számformátumot a felhasználó regionális beállításai szerint jeleníti meg, így magyar Excelben tizedesvessző látható.

## Kimeneti fájlnevek

Egy üzem esetén:

```text
<uzemkod>_ossszel_vetett_terulet.xlsx
```

Több üzem esetén:

```text
celzott_export.zip
```

A ZIP-ben minden üzemhez külön `.xlsx` fájl tartozik.

## Aktív export: Készletek – 5C és 6B

### Forrás és párosítás

- 5C forrásmunkalap: `t5_c`;
- 6B forrásmunkalap: `t6_b`;
- 5C záróérték: `osz = 5`;
- 6B zárókészlet: `osz = 12`;
- az elsődleges párosítás a szándékos `mXXXX` szabály alapján történik;
- a megnevezések mindkét forrásból megmaradnak második ellenőrzési adatként;
- összesítő sorok nem kerülnek az exportba;
- csak olyan részletes sor kerül bele, ahol az 5C vagy a 6B érték nem nulla.

A kódcsaládok megfeleltetése:

| 5C kódcsalád | 6B kódcsalád |
| --- | --- |
| `m55xx` | `m64xx` |
| `m56xx` | `m65xx` |
| `m57xx` | `m66xx` |
| `m58xx` | `m67xx` |

Ez nem egy soronként karbantartott párosítási lista. Az új, ugyanilyen kódszabályt követő sorokat a mapper automatikusan fel tudja venni.

### Kimeneti munkalap és oszlopok

A készlet-export ugyanabba az üzemenkénti Excel-fájlba kerül, mint a kiválasztott többi exporttárgy, külön `Készletek` munkalapon.

| Oszlop | Tartalom |
| --- | --- |
| `Partner azonosító` | kiválasztott üzemkód |
| `5C RowCode` | 5C sorazonosító |
| `5C megnevezés` | 5C sablon szerinti megnevezés |
| `5C záróérték` | saját készlet, 5C `osz = 5` |
| `6B RowCode` | 6B sorazonosító |
| `6B megnevezés` | 6B sablon szerinti megnevezés |
| `6B zárókészlet` | összes készlet, 6B `osz = 12` |
| `Saját készlet` | normál esetben a 5C záróértéke |
| `Vásárolt készlet` | 6B zárókészlet − 5C záróérték |
| `Ellenőrzés` | számítási állapot vagy hibaüzenet |

### Különleges esetek

- Ha a 5C és 6B érték is nulla, a sor kimarad.
- Ha a 5C érték nulla, a 6B érték pedig nem nulla, a különbözet teljes egészében vásárolt készlet.
- A jelenlegi sablonban a `m6510 – Vásárolt takarmánykeverék` csak a 6B-ben van. Ez bekerül, az 5C mezők üresek, a saját készlet `0,00`, a vásárolt készlet a teljes 6B érték.
- Ha a 6B érték kisebb az 5C értéknél, a két számított oszlop `0,00` lesz, az `Ellenőrzés` cella pedig piros `HIBA` jelzést kap.
- Ha a kódszabály alapján egy későbbi sablonban részletes 5C sorhoz nem található 6B-pár, az export hibával leáll, és nem készít bizonytalan párosítást.

### Excel-formázás

- az értékoszlopok valódi numerikus Excel-cellák;
- két tizedes megjelenítés `0.00` formátummal, amely magyar Excelben tizedesvesszőként jelenik meg;
- fejlécszűrő, rögzített fejléc és bekapcsolt rácsvonalak;
- vékony cellakeret és formázott fejléc;
- a hibás ellenőrzés piros, az érvényes állapot zöld háttérrel jelenik meg.

## Aktív export: Állatok – 6A

### Forrás és kiválasztási szabály

- Forrásmunkalap: `t6_a`;
- záróállomány-oszlop: `osz = 12` (`záróállomány (=)`);
- csak a nem nulla záróértékű sorok kerülnek exportba;
- az export nem korlátozott előre felsorolt fajokra, hanem a feltöltött 6A sablon minden nem összesítő sorát képes kezelni;
- a ló-, hal-, méh-, kecske-, nyúl- és egyéb állatsorok ezért ugyanúgy feldolgozhatók, mint a szarvasmarha, sertés, juh vagy baromfi sorai.

Az alábbi összesítő sorok nem kerülnek az exportba:

```text
m6209  LÓFÉLÉK ÖSSZESEN
m6239  SZARVASMARHA ÖSSZESEN
m6243  Hízósertés (50 kg felett) összesen
m6259  SERTÉS ÖSSZESEN
m6269  JUH ÖSSZESEN
m6309  BAROMFI ÖSSZESEN
```

Az összesítő sorokat a megnevezésben szereplő `összesen` vagy `mindösszesen` alapján szűrjük, ezért a későbbi, ugyanilyen jelölésű új sorok sem kerülnek bele automatikusan.

Az `előző sor t-ban` sorok nem összesítők. Ha a záróértékük nem nulla, bekerülnek, mert az előző állatcsoporthoz tartozó súlyadatot tartalmazzák.

### Kimeneti munkalap

```text
Állatok
```

Az állatos adatrész valódi Excel-táblaként készül, technikai táblaneve:

```text
tbl_Allatok
```

Ha az adott üzemben nincs nem nulla 6A záróérték, az `Állatok` munkalap csak a fejlécet tartalmazza, és nem kap `tbl_Allatok` objektumot. Így az Excel nem próbál üres táblázatot helyreállítani.

### Kimeneti oszlopok

| Oszlop | Tartalom |
| --- | --- |
| `Üzemkód` | a kiválasztott üzem kódja |
| `RowCode` | a 6A sorazonosítója, például `m6212`, `m6200` vagy `m6331` |
| `Sor megnevezése` | a 6A sablon szerinti megnevezés |
| `Mértékegység` | `db` állatlétszámhoz, `t` súlysorhoz, `család` a méhcsaládokhoz |
| `Záróérték` | numerikus 6A `osz = 12` záróállomány |

Példa:

| Üzemkód | RowCode | Sor megnevezése | Mértékegység | Záróérték |
| --- | --- | --- | --- | ---: |
| 20/4424 | m6212 | Üszőborjú 0,5 éves korig | db | 46,00 |
| 20/4424 | m6219 | Húshasznú tehén (nem fejt) | db | 166,00 |
| 20/4424 | m6224 | előző sor t-ban | t | 5,93 |
| 20/xxxx | m6200 | Lócsikó, 1 éves korig | db | 2,00 |
| 20/xxxx | m6331 | Pisztráng | db | 150,00 |

Az utolsó két sor szemléltető példa arra, hogy a ló- és halsorok is bekerülnek, ha az adott üzemben nem nulla záróértékük van.

### Excel-formázás

- a `RowCode` és az `Üzemkód` szöveges azonosítóként marad;
- a `Záróérték` valódi numerikus Excel-cellaként készül;
- a numerikus érték formátuma `0.00`, amely magyar Excelben tizedesvesszővel jelenik meg;
- fejlécszűrő, rögzített fejléc, bekapcsolt rácsvonal és vékony cellakeret készül;
- a sorok a 6A sablon sorrendjét követik, nem ábécésorrendben;
- az export ugyanabba az üzemenkénti munkafüzetbe kerül, mint a kiválasztott többi exporttárgy.

### Kimeneti fájlnevek

Ha csak az Állatok export van kiválasztva, egy üzem esetén:

```text
<uzemkod>_allatok.xlsx
```

Több üzem vagy több exporttárgy esetén az általános célzott export fájlnév- és ZIP-szabály érvényesül.
