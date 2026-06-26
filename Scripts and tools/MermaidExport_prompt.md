# VBA Macro — Mermaid flowchart export from Excel ("publication" sheet)

> Handover prompt. Give this to an AI or a developer to (re)build or continue the
> macro. It reflects every decision settled so far. The current working module is
> `MermaidExport.bas` (module name `MermaidExport`, entry point
> `Public Sub GenerateMermaidFiles()`).

---

## 0. Coding principles (these come first, do not drop them)

- **Keep the code as clean, compact, readable and simple as possible.** Prefer
  straightforward, literal logic over clever abstractions. No premature
  optimization. If a rule can be expressed plainly, express it plainly.
- All **code and comments in English**. **No accented or special characters in
  the VBA source** — build any accented *output* with `ChrW$()`.
- **Declare every variable outside loops** (use `ReDim` inside where needed) to
  avoid stale-value bugs.
- **Be robust against messy data**: the macro must never emit invalid Mermaid
  (no empty identifiers, no dangling links), even if cells are blank.
- Use **late binding** via `CreateObject` for `Scripting.Dictionary`,
  `Scripting.FileSystemObject`, and `ADODB.Stream` (no library references).

---

## 1. Context & entry conditions

- Reads the sheet named **`publication`** in the active workbook.
- **Always activate `publication` first.**
- Error + exit if: the sheet is missing, the workbook has never been saved
  (`wb.Path` empty), or there is no data.
- Row 1 is the header; data starts at row 2; no empty rows.

## 2. Input columns (sheet `publication`)

| Col | Meaning |
|-----|---------|
| A | ID |
| B | Thematique |
| C | Intitule |
| D | Palier (numeric, integer or decimal e.g. 0, 1, 1.5, 2) |
| E | Verticale |
| F–K | Description / Conseils / Justification / Public / Config / Exemple — **ignored** |
| L | Type de mesure: `M`, `M+`, `M-`, `R`, `R+`, `R-`, `R--` |
| M | Associated measure ID(s), comma-separated, may be empty |

Only **A, B, C, D, E, L, M** are used.

## 3. Output

Folder `"{WorkbookBaseName} - Export mermaid"` next to the workbook (created if
absent). All files `.mermaid`, **UTF-8 without BOM**.

1. `{base}.mermaid` — full export (all thematiques), HTML labels, includes
   inter-thematique links.
2. `{base}_excalidraw.mermaid` — full export in **plain text** (no HTML tags),
   for import into Excalidraw.
3. `{base}_{SafeThematique}.mermaid` — one per Thematique, HTML labels.
4. `{base}_{SafeVerticale}.mermaid` — one per Verticale, HTML labels, wrapped in
   its parent Thematique.

Partial exports (per Thematique / per Verticale) keep the full 3-level structure
**Thematique > Verticale > Palier** and omit inter-thematique links.

A single `BuildDoc(included() As Boolean, isFullExport As Boolean, plainText As
Boolean)` produces every file; the three booleans/filters select what goes in.

## 4. File structure (exact order)

1. **YAML config block**
   ```
   ---
   config:
     flowchart:
       defaultRenderer: "elk"
     look: classic
     theme: 'base'
     themeVariables:
       lineColor: '#284370'
       nodeBorder: '#284370'
       primaryTextColor: '#284370'
       primaryColor: '#E0E3E6'
       clusterBkg: '#D6C8DB'
       clusterBorder: '#6E1C73'
       titleColor: '#6E1C73'
   ---
   ```
2. **Global direction**: `flowchart LR`
3. **Seven `classDef` lines** (note: class name is `Pallier` with two L's):
   ```
   classDef Thematique font-size:14px,fill:transparent,stroke:#000000,color:#000000
   classDef Verticale font-size:26px,fill:#D6C8DB,stroke:#6E1C73,color:#6E1C73
   classDef Pallier font-size:20px,fill:#D6C8DB,stroke:#6E1C73,color:#284370
   classDef Mesure font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370
   classDef MesurePlus font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370
   classDef MesureMoins font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370
   classDef Recommendation font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370
   ```
4. **Node declarations** — every included node, declared once:
   - HTML mode: `  ID["<i>ID</i><br><b>[Type] {wrapped Intitule}</b>"]`
     where the Intitule is wrapped at word boundaries to lines ≤ 25 chars,
     joined with `<br>`, each line HTML-escaped.
   - Plain mode (Excalidraw): `  ID["ID - [Type] {plain Intitule}"]` — single
     line, no HTML tags.
5. **Subgraph structure** (see §5).
6. **Inter-palier links** — after each Verticale `end`.
7. **Inter-verticale links** — after each Thematique `end`.
8. **Inter-thematique links** — full export only, after all subgraphs.
9. **Class assignments** under `%% Apply style`.

## 5. Subgraph structure & layout

### Identifiers
- **Thematique ID** = `SafeID(name)`.
- **Verticale ID** = `{ThematiqueID}__{SafeID(name)}` — *scoped by Thematique*
  so identical verticale names under different thematiques never collide.
- **Palier ID** = `{VerticaleID}_Palier_{FmtPalier(value)}` (one L in `_Palier_`).
- `SafeID()`: transliterate French accents to ASCII, then replace every
  non-alphanumeric char with `_`. **Must never return an empty string** — fall
  back to a constant (e.g. `EMPTY`) for blank names.
- `FmtPalier()` = `Trim$(Str$(v))` → no trailing zeros, `.` decimal separator
  (palier IDs may contain a dot, e.g. `_Palier_1.5`; kept literal).

### Directions & titles
| Level | Direction | Title (HTML) | Title (plain) |
|-------|-----------|--------------|---------------|
| Thematique | `LR` | `<i>{FrenchMonth} {YYYY}</i><br><b>Dalton - Thematique {name}</b>` | `{FrenchMonth} {YYYY} - Dalton - Thematique {name}` |
| Verticale | `BT` | `<center>{wrap(name,30)}</center>` | `{name}` |
| Palier | see below | `Palier N` or `Palier N (vide)` | same |

`FrenchMonth`/`YYYY` use the current date.

### Palier gap-filling
Per Verticale: collect distinct Palier values, sort ascending. With
`minInt = floor(min)`, `maxInt = floor(max)`, insert an **empty** `Palier N
(vide)` for every integer in `[minInt, maxInt]` not present as an exact integer.
Non-integer paliers (e.g. 1.5) sort in place and do **not** mark their floor as
present. Empty paliers use `direction LR` and contain nothing.

### Measures inside a Palier — two cases
Order the palier's rows by type priority: `M+, M-, M, R+, R-, R--, R`
(stable sort; equal priority keeps data order).

- **If any measure of the palier is an arrow endpoint** → `direction LR`.
  Emit the arrows (§6), then plain node lines `  ID` for any measure not used in
  an arrow.
- **Otherwise (pure measures, no arrows)** → `direction TB`, laid out in **two
  columns**, row-major: left column = positions 1,3,5…, right = 2,4,6…. Chain
  each column top-down with the **invisible link** `~~~`. A column with a single
  node has no link, so emit that node as a plain line.
  *(This makes the two columns deterministic instead of relying on elk's
  auto-packing, which was unreliable.)*

## 6. Arrows

Driven **only** by column M. For every row with a non-empty M, split on commas;
for each entry emit `{Mvalue} --> {currentRowID}` (**source = column M value,
target = the current row**). Emit an arrow only when **both** endpoints are in
the included set (prevents referencing undeclared nodes in partial exports).
Any node that is an arrow source or target must **not** also appear as a plain
node line. Compute the set of arrow nodes in a global pre-pass.

## 7. Links between subgraphs

- **Inter-palier** (inside the Verticale, after its `end`):
  `%% Links between paliers for {vert}` then chain consecutive Palier IDs with
  `---`.
- **Inter-verticale** (after the Thematique `end`):
  `%% Links between verticales for {them}` then chain the Verticale IDs
  (first-appearance order, Thematique-scoped) with `---`.
- **Inter-thematique** (full export only, after all subgraphs):
  `%% Space thematiques` then chain Thematique IDs with `---`.

## 8. Class assignments (`%% Apply style`)

Only for elements present in the file: each Thematique → `Thematique`, each
Verticale → `Verticale`, each Palier → `Pallier`, each node by type
(`M`→`Mesure`, `M+`→`MesurePlus`, `M-`→`MesureMoins`, `R`/`R+`/`R-`/`R--`
→`Recommendation`, anything else → `Mesure`).

## 9. Text escaping

- **`EscapeHTML`** (HTML labels) — replace `&` **first**, then:
  `"`→`&quot;`, `<`→`&lt;`, `>`→`&gt;`, `'`→`&#39;`, `` ` ``→`&#96;`,
  `‘`/`’`(U+2018/2019)→`&#39;`, `«`/`»`→`&quot;`.
- **`EscapePlain`** (Excalidraw labels) — `"`→`'`, `<`→`(`, `>`→`)`; keep accents.

## 10. Encoding

UTF-8 **without BOM** via `ADODB.Stream`: write as text (Charset `UTF-8`),
re-open as binary, skip the 3 BOM bytes, save with overwrite. (BOM-stripping
matters: a BOM before the leading `---` breaks the YAML front-matter.)

## 11. Helper inventory (current module)

`GenerateMermaidFiles` (entry) · `BuildDoc` · `BuildPaliers` · `ToDouble` ·
`GetBaseName` · `OrderedUniqueArr` · `OrderedUniqueArrFiltered` · `FmtPalier` ·
`SamePalier` · `TypeToClass` · `TypePriority` · `FrenchMonth` · `EscapeHTML` ·
`EscapePlain` · `WrapTitle` · `SafeID` · `WriteUTF8NoBOM`.

## 12. Known caveats / open points

- **Two-column packing**: a palier with exactly **two** measures has no `~~~`
  link, so elk positions the two lone nodes itself (normally side by side).
  Three or more measures are pinned into columns deterministically.
- **Arrow direction** (source = column M, target = current row) reverses an
  earlier assumption — confirm it is intended for your data.
- **Decimal palier IDs** contain a `.` (e.g. `_Palier_1.5`); kept literal per
  spec — verify your renderer is happy with it.
- The `[Type]` shown in labels is column L; if it renders empty (`[] …`), column
  L is blank for those rows.

## 13. Tuning knobs (single values to change)

- Intitule wrap width: `WrapTitle(gIntitule(i), 25)`.
- Verticale title wrap width: `WrapTitle(vertName, 30)`.
- Number of measure columns: the `For c = 0 To 1` loop in the palier block.
- Thematique title size: `classDef Thematique font-size:14px,...`.
