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
- a Célzott export gomb a dialog mögött végig látható marad;
- a Bezárás gomb, a jobb felső `×` és az Esc billentyű bezárása is törli a dialog állapotát.
- ha egy kiválasztott exporttárgyhoz az adott üzemben nincs nem nulla adat, a munkalap fejlécével megmarad, de üres Excel-tábla és üres AutoFilter nem kerül rá;
- kivétel a `Támogatási jogcímek` lap: a teljes rögzített jogcímlistát akkor is kiírja, ha nincs forrásérték;
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

### Ősszel vetett terület fájlneve

Egy üzem esetén:

```text
<uzemkod>_ossszel_vetett_terulet.xlsx
```

Több üzem esetén:

```text
celzott_export.zip
```

A ZIP-ben minden üzemhez külön `.xlsx` fájl tartozik.

## Aktív export: Földterületi adatok – t1_a

### Forrás és kiválasztási szabály

- forrásmunkalap: `t1_a`;
- értékoszlopok: a sablon nem üres kódú `osz = 3–9` oszlopai;
- csak olyan sor kerül az exportba, amelyben legalább egy értékoszlop nem üres és nem nulla;
- az üres és nulla értékű cellák üresen maradnak, nem lesznek automatikusan nullává alakítva;
- a sorok a `t1_a` sablon sorrendjét követik;
- az összesítő és számított sorok is megmaradnak, ha van bennük adat;
- a sablon üres záró fejlécű technikai oszlopa nem kerül a kimenetbe.

### Kimeneti munkalap és oszlopok

A munkalap neve:

```text
Földterületi adatok
```

Az oszlopok a megjelenítő főnézetének sorrendjét követik:

| Oszlop | Tartalom |
| --- | --- |
| `FarmCode` | a kiválasztott üzem kódja |
| `RowCode` | a t1_a sablon sorazonosítója |
| `RowTitle` | a t1_a sablon sorának megnevezése |
| `saját tulajdon (+)` | t1_a `osz = 3` |
| `bérleti díjért bérbe adott terület (-)` | t1_a `osz = 4` |
| `ingyenesen bérbe adott terület (-)` | t1_a `osz = 5` |
| `bérleti díjért bérbe vett terület (+)` | t1_a `osz = 6` |
| `ingyenes bérbe vett terület (+)` | t1_a `osz = 7` |
| `üzem által használt összes terület (=)` | t1_a `osz = 8` |
| `földérték` | t1_a `osz = 9` |

Az `FarmCode`, `RowCode` és `RowTitle` technikai/azonosító mezők szövegként maradnak. A hét értékoszlop valódi numerikus Excel-cella, két tizedes formátummal. A munkalap fejlécszűrőt, rögzített fejlécet, rácsvonalat és vékony cellakeretet kap.

Ha egy üzemnél nincs érdemi t1_a adat, a `Földterületi adatok` munkalap csak fejléccel készül, üres Excel-tábla és üres AutoFilter nélkül.

### Kimeneti helye

A munkalap az üzemenkénti munkafüzet negyedik lapja, az `Ősszel vetett terület`, `Készletek` és `Állatok` munkalapok után, ha mindegyik exporttárgy ki van választva.

Ha csak ez az export van kiválasztva, egy üzem esetén a fájlnév:

```text
<uzemkod>_foldteruleti_adatok.xlsx
```

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

## Aktív export: Támogatási jogcímek – t7_c és t7_b1

### Forrás és értelmezés

- az AKG-, erdészeti és Natura 2000 erdőösszegek forrása a `t7_c`, `osz = 3`;
- a fiatal gazdálkodók CIS-YF összege a `t7_b1`, `osz = 4` (`eFt`); az `osz = 3` egységadat nem pénzösszeg, ezért nem kerül felhasználásra;
- a forrás pénzösszegei eFt-ban vannak, a kimenet Ft-ra váltja őket (`eFt × 1 000`);
- az éves területalapú támogatás és az állattenyésztési adatok nem részei ennek a lapnak;
- minden kiválasztott üzemben ugyanaz a fix jogcímlista szerepel, a forrásbeli üres és nulla értékektől függetlenül.

### Jogcím- és sorkódmapping

| Célűrlap-jogcím | FADN-sorkód | Forráslap / osz | Megjegyzés |
| --- | --- | --- | --- |
| Agrár-környezetgazdálkodási program | `m7094` | `t7_c / 3` | AKG összesítő |
| Szántó | `m7080` | `t7_c / 3` | Szántó összesítő |
| Horizontális szántó | `m70801` | `t7_c / 3` | Részjogcím |
| Horizontális szántó – Talajmegújító gazdálkodás szántó | `m70802` | `t7_c / 3` | A forrássor megnevezése strip tillre, sávos művelésre utal; a célűrlappal való megfeleltetés nem szó szerinti. |
| Szántó – Talajmegújító – no-till | `m70803` | `t7_c / 3` | Részjogcím |
| Natura 2000 szántó | `m70804` | `t7_c / 3` | Részjogcím |
| MTÉT Túzokvédelmi szántó | `m70805` | `t7_c / 3` | Részjogcím |
| MTÉT Madárvédelmi szántó | `m70806` | `t7_c / 3` | Részjogcím |
| MTÉT Kék vércse védelmi szántó | `m70807` | `t7_c / 3` | Részjogcím |
| Gyep | `m7082` | `t7_c / 3` | Gyep összesítő |
| Horizontális gyep | `m70821` | `t7_c / 3` | Részjogcím |
| MTÉT alföldi madárvédelmi | `m70822` | `t7_c / 3` | Részjogcím |
| MTÉT túzokvédelmi | `m70823` | `t7_c / 3` | Részjogcím |
| MTÉT hegy- és dombvidéki madárvédelmi | `m70824` | `t7_c / 3` | Részjogcím |
| MTÉT nappal lepke védelmi | `m70825` | `t7_c / 3` | Részjogcím |
| MTÉT gyeprezervátum | `m70826` | `t7_c / 3` | Részjogcím |
| Ültetvény | `m7084` | `t7_c / 3` | Ültetvény összesítő |
| Ültetvény – intenzív | `m70841` | `t7_c / 3` | Részjogcím |
| Ültetvény – extenzív | `m70842` | `t7_c / 3` | Részjogcím |
| Ültetvény – szőlő | `m70843` | `t7_c / 3` | Részjogcím |
| Horizontális nádas | `m7086` | `t7_c / 3` | Részjogcím |
| Erdőgazdálkodás támogatása | `m7215` | `t7_c / 3` | Mezőgazdasági területek első erdősítésére adott támogatás; külön forráskódsoron marad. |
| Erdőgazdálkodás támogatása | `m7216` | `t7_c / 3` | Egyéb erdészeti támogatások. |
| Natura 2000 erdő | `m7212` | `t7_c / 3` | Natura 2000 erdőre adott támogatás. |
| Kistermelői támogatási rendszer | – | – | Nem találtunk hozzá külön FADN-sorkódot; nem kap kitalált értéket vagy kódot. |
| Fiatal mezőgazdasági termelők támogatása (FIG) | `m7407` | `t7_b1 / 4` | Fiatal gazdálkodók támogatása (CIS-YF). |

Az `m7213` és `m7240` külön fiatalgazda-induló támogatási kódok, ezért nem olvadnak be automatikusan az `m7407` CIS-YF sorába. A célűrlap erdőgazdálkodási sorához tartozó `m7215` és `m7216` külön sorban marad, hogy a forráskódok ellenőrizhetők legyenek.

### Kimeneti munkalap és oszlopok

```text
Támogatási jogcímek
```

| Oszlop | Tartalom |
| --- | --- |
| `Üzemkód` | a kiválasztott üzem azonosítója |
| `Célűrlap jogcíme` | a célűrlap szerinti jogcím megnevezése |
| `FADN-sorkód` | a forrássablon sorazonosítója |
| `FADN-megnevezés` | a forrássablon eredeti sorneve |
| `Összeg (Ft)` | numerikus forintösszeg; forrás eFt érték × 1 000 |
| `Sor jellege` | `Összesítő` vagy `Jogcím`; az összesítők nem adhatók hozzá a részjogcímekhez |
| `Adatállapot` | forrásérték, nulla, üres forrásérték, hiányzó forrássor vagy hiányzó sorkód jelzése |

Az üres forrásértékhez numerikus `0` kerül, és az `Adatállapot` megkülönbözteti ezt a tényleges forrásbeli nullától. Ha a sablonban vagy a forrásadatban hiányzik a kód, az összeg cellája üres marad, és a státusz jelzi az okot. A kistermelői támogatás ezért szerepel a fix listában, de üres összeggel és `Nincs azonosított FADN-sorkód` állapottal.

Az AKG összesítő sorai (`m7094`, `m7080`, `m7082`, `m7084`) ellenőrzésre szolgálnak. Ne add őket hozzá a részjogcímek összegéhez. A lapon valódi numerikus összegek, egész Ft megjelenítés, fejlécszűrő, rögzített fejléc és az export többi lapjához illeszkedő formázás készül.

Ha mind az öt exporttárgy ki van jelölve, a támogatási lap az ötödik munkalap. Ha egy korábbi exporttárgy nincs kiválasztva, a lap ennek megfelelően előrébb kerül.

Egy üzem esetén, ha csak ez az export van kiválasztva:

```text
<uzemkod>_tamogatasi_jogcimek.xlsx
```

Több üzem vagy több exporttárgy esetén az általános célzott export fájlnév- és ZIP-szabály érvényesül.
