# CIB PDF -> XLSX Importer

.NET konzolos eszkoz CIB szamlakivonat PDF fajl feldolgozasahoz, es a tranzakciok exportjahoz XLSX formatumba.

A CLI a `System.CommandLine` csomagot hasznalja.

## Funkciok

- PDF beolvasas (`UglyToad.PdfPig`)
- Tranzakciok felismerese CIB kivonat szovegbol
- XLSX export (`ClosedXML`)
- Nyers PDF szoveg kulon munkalapon (`RawLines`) a hibakereseshez

## Kovetelmenyek

- .NET SDK 10.0+

## Telepites es futtatas

### 1. Restore es build

```bash
dotnet restore
dotnet build
```

### 2. Import futtatasa

```bash
dotnet run -- --input assets/cib.pdf --output assets/cib.xlsx
```

## CLI opciok

- `--input`, `-i` (kotelezo): bemeneti PDF fajl
- `--output`, `-o` (opcionalis): kimeneti XLSX fajl
- `--sheet-name`, `-s` (opcionalis): tranzakcios munkalap neve (alapertelmezett: `Transactions`)

Pelda:

```bash
dotnet run -- --input /utvonal/cib.pdf --output /utvonal/cib.xlsx --sheet-name "Kivonat"
```

## Kimeneti Excel szerkezet

- `Transactions` (vagy a megadott lapnev): feldolgozott tranzakciok
  - `Page`
  - `Date`
  - `Description`
  - `Amount`
  - `Balance` (jelenleg nem minden esetben toltott)
  - `RawLine`
- `RawLines`: oldalankenti nyers szovegblokkok a PDF-bol

## GitHub Actions release build

A projekt tartalmaz release workflow-t: `.github/workflows/release-build.yml`.

Trigger:

- `release.published`
- kezi futtatas: `workflow_dispatch`

Cel platformok:

- Linux: `linux-x64`
- Windows: `win-x64`
- macOS: `osx-x64`

A workflow self-contained, single-file buildet keszit, majd `.zip` vagy `.tar.gz` formatumban feltolti release assetkent.

## Ismert korlatok

- A parser jelenleg a tipikus CIB tranzakcios mintara optimalizalt.
- Elteto kivonat layout eseten finomhangolas szukseges lehet.

## Licenc

Belsos/egyedi hasznalatra keszult. Ha publikus repoba kerul, erdemes kulon `LICENSE` fajlt is hozzaadni.
