# 📋 Pastoral Social — Registry Lookup

A mobile app for looking up and managing the beneficiary registry of the Pastoral
Social (a parish community charity), built with **Streamlit**. It reads and writes
an `.xlsx` file hosted on **Google Drive**, which remains the source of truth.

The app's interface is in Portuguese; button labels, sheet names and column headers
are quoted below exactly as they appear in the app and in the spreadsheet.

---

## 🔍 Search

| Feature | Description |
|---|---|
| **Two search bars** | One for the record number, another for every other field |
| **Number search** | **Exact** match: searching "1" returns only record no. 1 |
| **General search** | Name, CPF, Reserva 1, CPF Reserva 1, Reserva 2 and CPF Reserva 2 |
| **OR logic** | Both bars can be used at once; the result is the union |
| **Partial search** | Works with fragments ("silva", "maria") |
| **Accent-insensitive** | "Antonio" finds "Antônio" and vice versa |
| **Ignores CPF punctuation** | "12345678909" finds "123.456.789-09" and vice versa |
| **Match highlighting** | Search terms are highlighted in yellow in the displayed fields |
| **Result priority** | Matches on Name or CPF come first |

In a number search, only the number in the card header is highlighted, so it doesn't
suggest the record matched on another field. Children's names are not highlighted,
because the children column is not a searchable field.

> CPF is the Brazilian individual taxpayer ID. *Reserva 1* and *Reserva 2* are the
> alternate people authorized to collect a beneficiary's food basket.

---

## 📇 Record display

| Feature | Description |
|---|---|
| **Card header** | Name, number and CPF |
| **Tier 1 fields** | Tipo, Status, CID 2026, Reserva 1, CPF Res. 1, Reserva 2, CPF Res. 2 |
| **Alerta para a Mesa** (front-desk alert) | In red and in a slightly larger font, when present |
| **Delivery history** | Compact 6-column grid; green cell for months delivered, grey for the rest |
| **Children** | Chips with an emoji by sex (👧 / 👦 / 🧒), name and age; shown only when there are children |
| **Dark mode** | Its own dark palette, search fields included |
| **Mobile-first** | Layout optimized for small screens |

---

## ✏️ Editing

### Available to everyone

| Feature | Description |
|---|---|
| **Mark this month's delivery** | Writes "OK" in the current month, after a confirmation popup. Disabled once delivered |
| **Cadastrar Criança** (add child) | Name, sex, age and a note. Shown when the record has no children |
| **Editar Crianças (N)** (edit children) | Loads the registered children to edit, remove or add more, all in a single save |

### Admin mode only

| Feature | Description |
|---|---|
| **Editar Dados** (edit details) | Tipo and Status (drop-down lists), CID 2026, Reserva 1 and 2, the alternates' CPFs and Alerta para a Mesa |

Only the fields that actually changed are written, and the Save button stays disabled
while nothing has changed.

---

## 🧺 Extra baskets (admin only)

| Feature | Description |
|---|---|
| **Dar Cesta Extra** (give an extra basket) | Sequential number pre-filled and editable; Number, Name and Reason are required |
| **Extra basket list** | A screen of its own with every record and counts of delivered and pending baskets |
| **Mark as delivered** | With a confirmation popup; writes to the ENTREGUE column |
| **Instant propagation** | A basket entered on one device shows up on the others within seconds, without re-reading Drive |
| **Sync indicator** | Shows whether saves are in progress and the time of the last sync |
| **Reload from Drive** | Pulls in changes made directly in the spreadsheet |
| **Failure recovery** | If a save fails, the data stays visible and there is a button to retry |

Extra baskets live on the `Sheet2` tab and are independent of the main registry.

---

## 🔐 Admin mode

A discreet link in the footer asks for the password, which is read from the app's
*Secrets*. Until a password is configured, admin access is unavailable rather than
open. Leaving admin mode closes any restricted screens automatically.

It's worth being clear about what this protects: it organizes who does what, but it
is not a security barrier. The password lives in the app process and there is no
limit on attempts. What really protects the data are the Drive service-account
credentials, which stay on the server only.

---

## 💾 Saving and data integrity

Every change goes to the `.xlsx` on Google Drive through a service account. Because
each save re-uploads the whole file, the app uses a few safeguards:

- **Serialized writes**: a single lock in the process ensures two simultaneous saves
  don't overwrite each other.
- **Content-based addressing**: the row is found by NUMERO and the column by its
  header name, always in the freshly downloaded file. Fixed positions are never used,
  because a row or column inserted into the spreadsheet would shift everything after it.
- **Refusal when ambiguous**: if the number isn't found, or more than one record has
  the same number, nothing is changed and the reason is shown.
- **Optimistic concurrency for children**: since an edit replaces the whole list, the
  save is refused if someone else changed the list in the meantime.
- **Write before display**: the screen reflects a change only after Drive confirms it,
  so it never shows as saved something that failed.
- **Asynchronous extra baskets**: the click responds immediately and the write
  continues in the background, batching pending changes into a single trip to Drive.
- **Double timestamp**: each change records when it was saved in the app and when it
  reached the spreadsheet; the gap reveals network waits.

---

## 🗂️ `.xlsx` file layout

### Tab 1 — registry

Two header rows:

| Row | Content |
|---|---|
| Row 1 | Tags: `SEARCH`, `Tier 1` or `Tier 2` |
| Row 2 | Column names |
| Row 3+ | Record data |

**`SEARCH` fields** (searchable): NUMERO, CPF, NOME, RESERVA 1, CPF RESERVA 1,
RESERVA 2, CPF RESERVA 2

**`Tier 1` fields** (shown on the card): TIPO, CID 2026, STATUS,
ALERTA PARA A MESA and the monthly delivery columns

**`Tier 2` fields** (not shown): address, phone, internal notes and the app's three
internal columns

The monthly columns must be **real dates** in row 2. The app shows them as `Mmm/YY`
(e.g. `Mar/26`) and uses the current month to enable the delivery button.

### Internal columns

| Header | Content |
|---|---|
| `CRIANÇAS` | JSON with the record's list of children |
| `Salvo no App` | When the user confirmed the save |
| `Gravado na Planilha` | When the spreadsheet was actually written |

The app finds these three by header name and also accepts their previous names
(`Última Alteração`, `Alterado Por:`, `Motivo da Alteração`), so renaming the
spreadsheet columns and publishing the code don't have to happen at the same time.

### Children JSON format

```json
{"criancas":[{"NÚMERO":"15","NOME":"Ana","SEXO":"Feminino","IDADE":8,"OBS":""}]}
```

Extra fields in the JSON are preserved by edits made in the app, even though they
don't appear in the form. A cell with invalid content is treated as "no children"
for display, and saving is refused so it isn't overwritten.

### Tab 2 — extra baskets

Header in the first row: `NÚMERO`, `NOME`, `CPF`, `TELEFONE`, `MOTIVO`, `ENTREGUE`.
