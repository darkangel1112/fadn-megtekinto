# Üzemeltetési útmutató

## Környezeti előfeltételek

- Python 3.12 vagy azzal kompatibilis Python 3.x környezet;
- virtuális környezet javasolt;
- hálózati elérés csak a csomagok telepítéséhez szükséges;
- a bemeneti Excel-fájlokhoz megfelelő helyi hozzáférés.

A függőségek a [requirements.txt](../requirements.txt) fájlban vannak.

## Helyi indítás

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m streamlit run app.py --server.headless true --server.address 127.0.0.1 --server.port 8503
```

Ellenőrzés:

```powershell
(Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8503/_stcore/health).Content
```

Elvárt válasz:

```text
ok
```

## Systemd telepítés

A repóban található [fadn-megtekinto.service](../fadn-megtekinto.service) egy Linuxos systemd-példa.

Jelenlegi beállításai:

- felhasználó/csoport: `darkangel`;
- munkakönyvtár: `/home/darkangel/webapps/apps/FadnMegtekinto`;
- cím: `127.0.0.1`;
- port: `8502`;
- base path: `/fadn`;
- újraindítás hiba esetén: bekapcsolva.

A tényleges szerveren a felhasználót, munkakönyvtárat és virtuális környezetet a helyi telepítéshez kell igazítani.

Tipikus telepítési lépések:

```bash
sudo cp fadn-megtekinto.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now fadn-megtekinto
sudo systemctl status fadn-megtekinto
```

Az alkalmazás a konfiguráció alapján jellemzően ezen az útvonalon érhető el:

```text
http://<szerver>/fadn/
```

## Frissítés

Éles frissítés előtt:

1. ellenőrizd a Git-ágat és a munkafát;
2. győződj meg arról, hogy nincs benne érzékeny adatfájl;
3. húzd le a kívánt commitot;
4. frissítsd a függőségeket;
5. indítsd újra a szolgáltatást;
6. ellenőrizd a státuszt és az egészségügyi végpontot.

Példa:

```bash
git pull --ff-only origin main
/home/darkangel/webapps/apps/FadnMegtekinto/.venv/bin/pip install -r requirements.txt
sudo systemctl restart fadn-megtekinto
sudo systemctl status fadn-megtekinto
```

## Naplózás

A systemd-szolgáltatás a standard kimenetet és hibakimenetet a journalba írja:

```bash
journalctl -u fadn-megtekinto -f
```

Hibakereséshez érdemes ellenőrizni:

- fut-e a folyamat;
- foglalt-e a konfigurált port;
- helyes-e a `WorkingDirectory`;
- létezik-e az `ExecStart` alatt megadott Streamlit;
- telepítve vannak-e a `requirements.txt` szerinti csomagok;
- elérhető-e a `/_stcore/health` végpont.

## Gyakori hibák

### „Hiányzó oszlopok az ömlesztett fájlból”

Az ömlesztett fájl első munkalapján ellenőrizd az `akod`, `sor`, `osz`, `ertek`, `dimenzio1` és `dyn_row_serial` oszlopokat.

### Sok ismeretlen rekord

Az alkalmazás az ömlesztett `sor` kódokat a sablonból feloldott mappinggel kapcsolja munkalaphoz. Ha a sablon és az ömlesztett fájl nem ugyanahhoz az exportstruktúrához tartozik, sok kódhoz nem lesz munkalap.

### A célzott export gomb letiltva marad

Mindkét forrásfájlt fel kell tölteni, és az ömlesztett adatoknak tényleges rekordokat kell tartalmazniuk. Az 5C célzott exporthoz a sablonban a `t5_c` munkalapnak is szerepelnie kell.

### A böngészőben régi állapot látható

Ellenőrizd, hogy nem régi Streamlit-munkamenet maradt-e nyitva. A források törlése, majd újrafeltöltése új munkamenetet és új beolvasást eredményez.

## Biztonság

- Éles partneradatokat ne tölts fel nyilvános vagy megosztott fejlesztői környezetbe.
- Forrás- és teszt-Excel ne kerüljön Git-be.
- Ne kerüljön token, jelszó vagy egyéb titok a repóba.
- A `.streamlit/secrets.toml` ki van zárva a Git-ből, de a titkokat ennek ellenére külön titokkezeléssel kell kezelni.
