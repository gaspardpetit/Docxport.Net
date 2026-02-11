# Microsoft Word Field Code Syntax

## Introduction

Microsoft Word’s field code system is an **ad‑hoc mini‑language** for conditional logic, calculations, mail merges and dynamic document content.  The language evolved incrementally; there is **no formal grammar** or unified specification.  Field codes are encapsulated between curly braces (`{ }`) inserted with **Ctrl + F9** and are evaluated by Word when a document is opened, printed or when the user presses **F9** to update fields.  Because the syntax and evaluation rules are poorly documented and fragile, this document consolidates authoritative information from Microsoft documentation and trusted secondary sources to provide an exhaustive reference for implementing a **parser/interpreter** for Word fields.

Field codes consist of:

- A **field type** (`IF`, `SET`, `REF`, `SEQ`, `DATE`, `COMPARE`, etc.).
- An optional **expression or parameters** whose syntax depends on the field type.
- Optional **switches** beginning with a backslash (`\`) that modify the way results are formatted.

This specification is organised into conceptual sections with examples and references.  Each section lists applicable field types, syntax rules, supported parameters and switch behaviour.  A final section discusses implementation considerations for offline evaluation.


## Scope, Non-Goals, and Intentional Exclusions

This document describes Microsoft Word field code semantics at a level sufficient for correct parsing, evaluation, and integration outside of Word itself. Certain categories of Word behavior are intentionally out of scope, not due to incomplete research, but because they are intrinsically dependent on Word’s proprietary layout, style, or engine subsystems and cannot be faithfully reproduced without embedding or re-hosting Word.

This document does not attempt to fully specify or re-implement layout-driven and pagination-dependent fields whose results depend on Word’s final line breaking, pagination, widow and orphan control, and floating object placement. For these fields, evaluation semantics are described, but exact numeric results are only defined once pagination is stable.

Style-engine and list-engine constructs whose behavior is governed by Word’s internal numbering tables, abstract list definitions, and paragraph style inheritance rather than explicit field parameters are also excluded from full specification. This includes fields such as LISTNUM and numbering produced implicitly by list styles rather than by SEQ or related fields.

Legacy automatic numbering fields (AUTONUM, AUTONUMOUT, AUTONUMLGL) are described only at a high level. Their behavior is driven by undocumented, version-specific layout logic and is therefore non-deterministic outside Word.

Bibliography and citation fields (CITATION, BIBLIOGRAPHY) are documented in terms of construction, inputs, and switches, but not fully re-implemented. Final formatting depends on Word’s internal citation-style engine and locale-aware rendering logic, which are not publicly specified.

UI-driven constructs whose configuration is captured through dialogs rather than explicit field code syntax (for example AddressBlock and GreetingLine layout choices) are described in terms of evaluation behavior rather than UI configuration mechanics.

Finally, this document does not enumerate version-specific bugs or historical inconsistencies across Word releases. Where behavior is known to vary, the stable semantic model is described instead of cataloging defects.

These exclusions are deliberate. The intent is to provide a sound, defensible, and implementable semantic model for all declarative and data-driven field behaviors, while explicitly identifying the boundaries where Word’s internal engines become authoritative.

---

## Using and Updating Field Codes in Microsoft Word

Microsoft Word field codes are not plain text constructs; they are structured document objects with explicit insertion, display, update, locking, and navigation semantics. Correct use and interpretation of fields requires understanding the keyboard-level operations Word exposes for interacting with them, as these operations directly affect evaluation state, visibility, and persistence.

Field braces must be inserted using **Ctrl + F9**; typing literal `{` and `}` characters does not create a field and will not be recognized by Word’s field engine. Once inserted, a field exists as a discrete object that stores both its *code* and its *most recently computed result*. Word allows independent control over whether the user sees the code or the result, and whether the field is eligible for recomputation.

Field evaluation is triggered explicitly rather than continuously. Pressing **F9** updates the currently selected field or selection range. Selecting the entire document (Ctrl + A) followed by F9 updates all fields in the current story (typically the main document body). Headers, footers, text boxes, shapes, and other stories are not reliably updated by a single global select-all operation and may require separate updates or print-time refresh. Display toggles such as **Alt + F9** (all fields) and **Shift + F9** (current field) switch between code view and result view only; they do not trigger evaluation.

Word also supports explicit field state manipulation. Fields can be locked using **Ctrl + F11**, preventing them from updating while preserving their last computed result. Locked fields still participate in nested evaluations as literal values. Fields can be permanently unlinked using **Ctrl + Shift + F9**, which replaces the field with its current result text and removes all future update capability. Navigation between fields is provided by **F11** and **Shift + F11**, allowing sequential traversal of field objects independent of document text.

From an implementation perspective, these user-facing controls correspond directly to internal state: whether a field is locked, whether it is dirty (out of date), whether its code or result is displayed, and whether it remains linked to an evaluation mechanism. Any Word-compatible interpreter or document processor must model these states explicitly to match observed Word behavior.

### Field Keyboard Commands (Summary)

| Shortcut | Effect |
|--------|--------|
| **Ctrl + F9** | Insert a new field with proper field braces |
| **F9** | Update the selected field(s) |
| **Ctrl + A → F9** | Update all fields in the current story |
| **Alt + F9** | Toggle display of all field codes vs. results |
| **Shift + F9** | Toggle display of the selected field’s code vs. result |
| **Ctrl + Shift + F9** | Unlink field (convert to literal text) |
| **Ctrl + F11** | Lock field (suppress updates) |
| **Ctrl + Shift + F11** | Unlock field |
| **F11** | Jump to next field |
| **Shift + F11** | Jump to previous field |

### References

Microsoft Support – Update fields  
https://support.microsoft.com/en-us/office/update-fields-7339a049-cb0d-4d5a-8679-97c20c643d4e

Microsoft Support – Keyboard shortcuts in Word  
https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-word-95ef89dd-7142-4b50-afb2-f762f663ceb2

Graham Mayor – Formatting Word fields  
https://www.gmayor.com/formatting_word_fields.htm


## 1 Fundamentals and Evaluation Semantics

### 1.1 Inserting and toggling fields

- Field codes **must be inserted** using **Ctrl + F9**; typing literal braces will not create a field.  Word toggles between displaying the code and the result with **Alt + F9** and updates the selected field with **F9**([Formatting Word fields (Graham Mayor)](https://www.gmayor.com/formatting_word_fields.htm)).  When toggled, Word shows either the raw code (`{ … }`) or the evaluated text.

- The **Field dialog** (Insert → Quick Parts → Field) inserts most fields and includes basic formatting options.  For advanced formatting, users edit the code directly in the document([Formatting Word fields (Graham Mayor)](https://www.gmayor.com/formatting_word_fields.htm)).

- **Field results are transient.**  When a field updates, any manual formatting on the result is lost unless the `\* MERGEFORMAT` switch is used([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  The `\* MERGEFORMAT` switch applies previous result formatting to the new result([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

### 1.2 Strings vs numbers and operators

- Word does **not** have explicit typing.  Expressions are treated as strings unless a numeric comparison or arithmetic forces numeric coercion.  To ensure numeric behaviour, avoid enclosing numbers in quotes and use numeric operators.

- Comparison operators supported by `IF`, `COMPARE` and `SKIPIF` fields are:
  - `=` (equal), `<>` (not equal), `>` (greater than), `<` (less than), `>=` (greater or equal) and `<=` (less or equal)([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).
  - Operators must be surrounded by spaces([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).

- **Wildcards:** In `IF` or `SKIPIF` comparisons, the second expression may contain `?` (any single character) or `*` (any string).  When using `*`, the portion of the first expression that matches the asterisk plus the remaining characters in the second expression cannot exceed 128 characters([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).

- **Quoting:** Expressions containing spaces must be enclosed in quotation marks([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).  In numeric picture switches, simple formats with no spaces do not require quotes, but complex formats or those containing spaces or text must be enclosed in quotes([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

### 1.3 Nested fields and evaluation order

- Word evaluates nested fields **inside‑out**.  When a field contains other fields (e.g., `{ IF { DOCVARIABLE Status } = "Approved" "Yes" "No" }`), the innermost fields are evaluated first.

- Field results can be fed into other fields.  For example, `SET` assigns a value which `REF` later reads.

- When building an interpreter, maintain a call stack to evaluate nested fields and update variables/bookmarks accordingly.

### 1.4 Regional settings

- Numeric and date/time formatting uses the **decimal symbol** and **digit grouping symbol** defined in the system’s regional settings([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Date fields also use locale‑specific month and day names.

---

## 2 General Format Switches

Switches modify how Word displays field results.  They begin with `\` followed by a character identifying the switch type.  The three general switch classes are **format switch (\*)**, **numeric format switch (\#)** and **date‑time format switch (\@)**([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

### 2.1 Format switch (\*)

The `\*` switch defines textual and numeric transformations and retains formatting when fields update.  It accepts one of the following named formats:

#### 2.1.1 Capitalisation formats

| Switch | Description | Example |
|---|---|---|
| `\* Caps` | Capitalises the first letter of each word([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ FILLIN "Type your name:" \* Caps }` displays `Luis Alverca` even when entered in lowercase([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* FirstCap` | Capitalises only the first letter of the first word([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ COMMENTS \* FirstCap }` displays `Weekly report on sales`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* Upper` | Converts the result to uppercase([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ QUOTE "word" \* Upper }` displays `WORD`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* Lower` | Converts the result to lowercase([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ FILENAME \* Lower }` shows a file name in all lower‑case([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |

#### 2.1.2 Number formats

The following named formats convert numbers into other representations.  When used, the case of the word (e.g., `ALPHABETIC` vs `alphabetic`) determines the case of the result([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  These formats are useful with `SEQ`, `PAGE`, and numeric `= (Formula)` fields.

| Switch | Description | Example |
|---|---|---|
| `\* alphabetic` | Displays numbers as alphabetic characters.  The case of `alphabetic` controls the case of the output([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ SEQ appendix \* ALPHABETIC }` displays `B`; `{ SEQ appendix \* alphabetic }` displays `b`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* Arabic` | Displays numbers as Arabic numerals([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ PAGE \* Arabic }` shows `31`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  A variant `ArabicDash` inserts hyphens around the number([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* CardText` | Converts numbers to cardinal text (words).  Combined with another `\*` switch to control capitalisation([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ = SUM(A1:B2) \* CardText }` yields `seven hundred ninety`; adding `\* Caps` makes `Seven Hundred Ninety`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* DollarText` | Similar to `CardText` but inserts `and` at the decimal and expresses cents as a fraction of 100([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ = 9.20 + 5.35 \* DollarText \* Upper }` becomes `FOURTEEN AND 55/100`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* Hex` | Displays numeric results as hexadecimal([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  | `{ QUOTE "458" \* Hex }` outputs `1CA`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* OrdText` | Displays ordinals as words([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ DATE \@ "d" \* OrdText }` yields `twenty‑first`; adding `\* FirstCap` produces `Twenty‑first`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* Ordinal` | Displays ordinals as Arabic numerals followed by `st`, `nd`, `rd` or `th`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ DATE \@ "d" \* Ordinal }` yields `30th`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* roman` | Displays numbers as Roman numerals.  The case of `roman` sets the case of the output([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ SEQ CHAPTER \* roman }` yields `xi`; `{ SEQ CHAPTER \* ROMAN }` yields `XI`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |

#### 2.1.3 Character formatting

| Switch | Description | Example |
|---|---|---|
| `\* Charformat` | Applies the formatting of the first letter of the field code to the entire result([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | If the `R` in `{ REF chapter2_title \* Charformat }` is bold, the entire referenced title appears bold([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `\* MERGEFORMAT` | Preserves the manual formatting applied to the previous field result when the field updates([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Word inserts this switch by default when fields are created through the Field dialog([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |

**Implementation note (DocxportNet):**
- **CHARFORMAT**: when synthesising a result, copy run properties from the *first character of the field code* (the first field‑code run) and apply them to the entire result run.
- **MERGEFORMAT**: when synthesising a result, reuse the *cached result run properties* (per‑run) from the previous field result.  If no cached result exists, fall back to an unstyled run.

This reflects Word’s runtime behaviour: CHARFORMAT ties result styling to the field code, while MERGEFORMAT preserves prior result styling.  MERGEFORMAT is the only case where evaluation depends on cached result formatting.

### 2.2 Numeric format switch (\#)

The `\#` switch, used with Formula (`=`), `SET`, `REF` and other numeric fields, defines custom numeric formatting.  Word constructs the numeric picture by combining format items([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)):

- **0 (zero):** forces display of a digit; if no digit exists, `0` is inserted([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ = 4 + 5 \# 00.00 }` outputs `09.00`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **# (pound):** digit placeholder; if no digit exists, a space is displayed([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ = 9 + 6 \# $### }` displays `$ 15`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **x:** drops digits to the left of the `x`.  An `x` to the right of the decimal rounds the result to that place([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Examples: `{ = 111053 + 111439 \# x## }` yields `492`; `{ = 1/8 \# 0.00x }` yields `0.125`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **`.` (decimal point):** determines decimal position([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **`,` (digit grouping symbol):** inserts grouping separators (e.g., thousands) according to regional settings([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **`-` (minus sign):** prepends a minus sign for negative numbers; inserts a space if the number is positive or zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **`+` (plus sign):** prepends a plus sign for positive numbers, a minus sign for negatives, or a space for zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **Other characters (`%`, `$`, `*`, etc.):** inserted as literal characters in the formatted result([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ = netprofit \# "##%" }` produces `33%`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **Positive/negative/zero sections:** a format of the form `"positive; negative"` or `"positive; negative; zero"` specifies separate formats for each case([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ Sales95 \# "$#,##0.00;-$#,##0.00" }` shows negative values with a minus sign and positive values normally([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **'text':** encloses literal text to be inserted([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ = Price * 8.1% \# "$##0.00 'is sales tax'" }` yields `$347.44 is sales tax`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **`numbereditem`:** includes the number of the most recent item numbered with the Caption command or a `SEQ` field.  The item identifier (e.g., "table") must be enclosed in grave accents (`)([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

### 2.3 Date‑time format switch (\@)

Used with date/time fields (`DATE`, `TIME`, `PRINTDATE`, `CREATEDATE`, etc.), `\@` specifies a picture for formatting the date and time([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Combine the following items to build custom patterns:

#### 2.3.1 Date instructions([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))

| Token | Meaning | Example |
|---|---|---|
| `M` | Month number (1 – 12) without leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | July → `7`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `MM` | Month number with leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | July → `07`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `MMM` | Abbreviated month name([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | July → `Jul`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `MMMM` | Full month name([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `d` | Day number without leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | 6 → `6`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `dd` | Day number with leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | 6 → `06`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `ddd` | Abbreviated weekday name([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | Tuesday → `Tue`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `dddd` | Full weekday name([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `yy` | Two‑digit year([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | 1999 → `99`, 2006 → `06`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `yyyy` | Four‑digit year([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |

#### 2.3.2 Time instructions([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))

| Token | Meaning | Example |
|---|---|---|
| `h` or `H` | Hour without leading zero (12‑hour `h` or 24‑hour `H`)([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `9 AM` → `9`, `5 PM` → `17` when using `H`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `hh` or `HH` | Hour with leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `9 AM` → `09`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |
| `m` | Minutes without leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `mm` | Minutes with leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `s` | Seconds without leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `ss` | Seconds with leading zero([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). |
| `AM/PM` or `am/pm` | Displays `A.M.`/`P.M.` in uppercase or lowercase([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)). | `{ TIME \@ "h AM/PM" }` displays `9 AM`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)) |

#### 2.3.3 Additional formatting([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))

- **'text':** encloses literal text to insert within a date/time pattern([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ TIME \@ "HH:mm 'Greenwich mean time'" }` produces `12:45 Greenwich mean time`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **Characters:** characters such as `:`, `-`, `*` or spaces can be inserted directly([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ DATE \@ "HH:mm MMM-d, 'yy" }` yields `11:15 Nov-6, '99`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **`numbereditem`:** includes the number of the most recent item numbered by a Caption or `SEQ` field within a date/time format([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).  Example: `{ PRINTDATE \@ "'Table' `table` 'was printed on' M/d/yy" }` might display `Table 2 was printed on 9/25/02`([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

Note: simple date/time formats without spaces or text do not require quotation marks, but complex patterns must be quoted([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

---

## 3 Formula and Calculation Fields

### 3.1 Formula field (`=`)

The Formula field performs arithmetic operations or evaluates functions within tables or documents.  Its general syntax is `{ = Expression [ \# NumericFormat ] }`.  Key aspects:

- **Expressions:** may include numbers, cell references (e.g., `A1`, `B3`), bookmarks, nested fields or functions([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm))([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).
- **Operators:** addition (`+`), subtraction (`-`), multiplication (`*`), division (`/`), exponentiation (`^`), and percentage (`%`)([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).
- **Comparison operators:** the same as for `IF` (`=`, `<>`, `>`, `<`, `>=`, `<=`) can be used inside `=` to produce a Boolean (1 or 0) result([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).
- **Functions:** Word supports functions such as `ABS`, `AND`, `AVERAGE`, `COUNT`, `MAX`, `MIN`, `MOD`, `PRODUCT`, `ROUND`, `SIGN`, `SUM`, etc. (complete list in documentation)([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).  Functions operate on cell ranges, bookmarks or numeric values.
- **Referencing table cells:** use A1‑style references; row numbers start at 1 and column letters at A([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).  Use `ABOVE`, `LEFT`, `RIGHT`, and `BELOW` as range specifiers.
- **Numeric formatting:** apply the `\#` switch to specify how the result appears([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).  Example: `{ = SUM(ABOVE) \# $,0.00 }` outputs a currency value with two decimal places([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).


Below is an exhaustive list of functions supported by Word’s formula field (the “=(Formula)” field) as documented in the official Word 2003 help. These functions can be used in field expressions to perform calculations or logical evaluations. Functions with empty parentheses (AVERAGE(), COUNT(), MAX(), MIN(), PRODUCT(), SUM()) accept any number of arguments and can also take table‑cell references (e.g., A1:B3). Arguments are separated by commas or semicolons depending on your regional settings

| Function         | Purpose (brief)                                                                                | Notes                                                    |
| ---------------- | ---------------------------------------------------------------------------------------------- | -------------------------------------------------------- |
| **`ABS(x)`**     | Returns the positive value of a number or formula regardless of its sign                       | Single argument                                          |
| **`AND(x,y)`**   | Returns `1` if both logical expressions are true, otherwise `0`                                | Takes two arguments                                      |
| **`AVERAGE()`**  | Calculates the average of the provided values or cell references                               | Accepts a list of values or cell references              |
| **`COUNT()`**    | Returns the number of items in the list of arguments                                           | Accepts a list of values or cell references              |
| **`DEFINED(x)`** | Returns `1` if the expression `x` is valid and evaluates without error; returns `0` otherwise  | Useful for testing whether a bookmark or variable exists |
| **`FALSE`**      | Always returns `0`                                                                             | Takes no arguments                                       |
| **`INT(x)`**     | Rounds down `x` to the nearest whole number (removes fractional part)                          | Single argument                                          |
| **`MIN()`**      | Returns the smallest value in the list of arguments                                            | Accepts a list of values or cell references              |
| **`MAX()`**      | Returns the largest value in the list of arguments                                             | Accepts a list of values or cell references              |
| **`MOD(x,y)`**   | Returns the remainder after dividing `x` by `y`                                                | Requires two numeric arguments                           |
| **`NOT(x)`**     | Returns `0` if logical expression `x` is true; returns `1` if `x` is false                     | Single argument (often used inside IF conditions)        |
| **`OR(x,y)`**    | Returns `1` if either or both logical expressions are true; returns `0` only if both are false | Takes two arguments                                      |
| **`PRODUCT()`**  | Multiplies all the values or expressions provided                                              | Accepts a list of values or cell references              |
| **`ROUND(x,y)`** | Rounds the value `x` to `y` decimal places (can be negative to round to tens/hundreds)         | `x` must be numeric; `y` must be an integer              |
| **`SIGN(x)`**    | Returns `1` if `x` is positive; returns `–1` if `x` is negative                                | Single argument                                          |
| **`SUM()`**      | Adds all the values or expressions provided                                                    | Accepts a list of values or cell references              |
| **`TRUE`**       | Always returns `1`                                                                             | Takes no arguments                                       |

Usage notes and examples

- Multiple arguments: Functions with empty parentheses can take several values or cell references. For example, `=SUM(A1, A2, B3:B5)` or `=AVERAGE(LEFT)` (the `LEFT` keyword averages cells to the left of the formula cell).

- Logical functions: `AND`, `OR`, `NOT`, `TRUE` and `FALSE` return `1` (true) or `0` (false), allowing you to build conditional expressions. In formulas, you can combine them with relational operators (`=`, `<`, `>`, etc.) and use the result in an `IF` field.

- Rounding: `INT(x)` truncates decimals; `ROUND(x,y)` gives more flexible rounding (e.g., `=ROUND(123.456,2)` → `123.46`).

- Testing definitions: `DEFINED(x)` is handy for checking whether a bookmark or variable has been set (returns 1 if it exists).

These functions are supported in Word 2003 and continue to be valid in later versions of Word for formula fields. There are no additional built‑in arithmetic or trigonometric functions beyond those listed above.

#### New or changed functions since Word 2003

| Function               | Earliest version found                                                                                       | Purpose and behaviour                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               | Evidence                                                                                                                                                                                                                                                                                   |
| ---------------------- | ------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **`IF()`**             | Word 2013 (function appears in the Word 2013/2016/2019/365 UI; it is absent from the Word 2003 formula list) | Evaluates a logical test and returns one value if the test is true and another if it is false.  The function requires three arguments: `IF(test, value_if_true, value_if_false)`.  For example, `=IF(SUM(LEFT)>=10,10,0)` returns 10 if the sum of values to the left of the formula is at least 10, otherwise 0.  Official Word 2019 documentation lists `IF()` among the available functions, and a 2021 training article covering Word 2013/2016/2019/365 describes using `=IF(test,true,false)` in Word tables. | Microsoft’s current support page for Word formulas lists `IF()` as an available function and explains that it evaluates a test and returns the second or third argument depending on whether the test is true.  A Word training guide (applicable to Word 2013‑365) shows the same syntax. |
| **`TRUE()` (revised)** | Word 2013 (change from a constant to a function)                                                             | In Word 2003 the constant `TRUE` simply returns 1.  In later versions it appears as a function `TRUE(argument)` that evaluates a logical expression; it returns 1 if the argument is true and 0 if it is false.  This allows you to evaluate a boolean expression directly inside a formula.                                                                                                                                                                                                                        | The Word 2019 support page shows `TRUE()` as taking one argument and returning 1 if the argument is true.  The Word 2003 documentation lists only a constant `TRUE` that returns 1.                                                                                                        |




### 3.2 Set and use variables in calculations

Variables can be stored in bookmarks using the `SET` field and later referenced using `REF` or within formulas.  Example (from Microsoft’s Set field documentation):

```text
{ SET UnitCost 25 }
{ SET Quantity { FILLIN "Enter number of items ordered:" } }
{ SET SalesTax 10% }
{ SET TotalCost { = (UnitCost * Quantity) + ((UnitCost * Quantity) * SalesTax) } }
This confirms your order of our book. You ordered { REF Quantity } copies at { REF UnitCost \# "$#0.00" } apiece. Including sales tax, the total comes to { REF TotalCost \# "$#0.00" }([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).
```

In this example, numeric results are formatted using the numeric picture switch.

---

## 4 Conditional and Comparison Fields

### 4.1 IF field

**Syntax:**

```text
{ IF Expression1 Operator Expression2 TrueText [ FalseText ] }
```

- **Expression1 / Expression2:** may be bookmark names, strings, numbers, nested fields or formulas.  Strings containing spaces must be enclosed in quotes([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).  When you want to compare strings to literal text, enclose the second expression in quotes([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).

- **Operator:** one of `=`, `<>`, `>`, `<`, `>=`, `<=`([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).  Operators must be surrounded by spaces.

- **TrueText / FalseText:** the text (or fields) displayed when the comparison evaluates to true or false.  If `FalseText` is omitted, nothing is inserted when the condition is false([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).

- **Wildcards:** `?` (matches one character) and `*` (matches zero or more characters) are permitted in Expression2 when using `=` or `<>`([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).

- **Nested field errors:** Word treats the *result text* of nested fields as the literal value used in comparisons.  For example, if `{ REF Missing }` yields `Error! Reference source not found.`, then `{ IF { REF Missing } = "" "Empty" "Not Empty" }` evaluates to `"Not Empty"` because the left expression is the error string, not an empty value.  This behaviour is observed in Word, but is not explicitly specified in Microsoft’s field code documentation.  
  **DocxportNet note:** In our evaluation, missing `REF` inside an `IF` currently uses the `REF` error text (not the raw field keyword). Some Word builds appear to display just `REF` for this case; we treat that as a documented divergence.

- **Examples:**

  *Simple numeric comparison:* `{ IF Order >= 100 "Thanks" "The minimum order is 100 units" }` displays `Thanks` when the order bookmark is ≥ 100([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).

  *Comparing merge fields and inserting text:* `{ IF { MERGEFIELD State } = "NY" "New York address details" "" }` uses a merge field inside the condition.

  *Nested logic:* complex conditions can be built by nesting `IF` or using `AND`/`OR` functions inside a `=` field; e.g., `{ IF { =AND({MERGEFIELD Rate}>100,{MERGEFIELD Discount}>0.2) } = 1 "Special discount applies" "Regular pricing" }`.  Additional examples appear in the 2003 documentation([Examples of IF fields (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefExamplesOfIFFields1.htm)).

### 4.2 COMPARE field

The `COMPARE` field compares two values and returns `1` if the comparison is true or `0` otherwise([COMPARE field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-compare-field-60bfb300-c58d-4f2f-8255-f1a9707390c8)).  Its syntax mirrors the IF field without TrueText/FalseText:

```text
{ COMPARE Expression1 Operator Expression2 }
```

Expressions and operators follow the same rules as `IF`([COMPARE field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-compare-field-60bfb300-c58d-4f2f-8255-f1a9707390c8)).  Use `COMPARE` when you need a numeric Boolean result (1/0) to feed into other fields (e.g., formulas or nested `IF`s).

### 4.3 SKIPIF field

`SKIPIF` is designed for mail merge: if the comparison is true, the current record is skipped and Word proceeds to the next record; otherwise the document is merged normally([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).  Syntax:

```text
{ SKIPIF Expression1 Operator Expression2 }
```

- Use the same operators and expression rules as the IF field([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).
- Do **not** combine `SKIPIF` with a `NEXT` field([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).
- Example: `{ SKIPIF { MERGEFIELD Order } < 100 }` prevents a merged document from being created when the `Order` value is less than 100([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).

### 4.4 NextIf and SkipIf in mail merges (brief mention)

`NEXTIF` is similar to `SKIPIF` but forces Word to skip to the next record if the condition is true; it isn’t documented as extensively and is rarely used.  When building an interpreter, treat `NEXTIF` similarly to `SKIPIF` by evaluating the condition and advancing the data cursor accordingly.

---

## 5 Data and Variable Fields

### 5.1 SET field

The `SET` field assigns a value to a **bookmark** (variable).  Syntax:

```text
{ SET Bookmark "Text" }
```

- `Bookmark` is the variable name; it behaves like a bookmark in the document([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).
- `"Text"` is the value assigned.  Text must be enclosed in quotes; numbers need not be([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).  The value may be the result of another field([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).
- To display the variable, insert a `REF` field referencing the bookmark([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).
- Example: the order‑form example above shows how `SET` fields create variables (`UnitCost`, `Quantity`, `SalesTax`, `TotalCost`) and `REF` fields display them([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).

### 5.2 REF field

`REF` displays the value of a bookmark or other referenced item.  Syntax:

```text
{ REF Bookmark [ \Switches ] }
```

Word supports eight field‑specific switches for REF plus a set of general format switches that apply to any field:


| Switch | Behaviour                                                                                                                                                                                                                                                                     | Sources                            |
| ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------- |
| `\d`   | **Separator characters.** Specifies the characters used to separate sequence numbers (such as chapter numbers) and page numbers.                                                                                                                                              | Word 2003 Help; Microsoft Support. |
| `\f`   | **Footnote/endnote insertion.** Increments footnote, endnote or annotation numbers marked by the bookmark and inserts the corresponding note text.  Example: `{ REF Note1 \f }` after footnote 2 displays the footnote reference mark “3” and inserts the text of footnote 1. | Word 2003 Help; Microsoft Support. |
| `\h`   | **Hyperlink.** Turns the field result into a hyperlink to the bookmarked paragraph.                                                                                                                                                                                           | Word 2003 Help; Microsoft Support. |
| `\n`   | **Paragraph number without trailing periods.** Displays the entire paragraph number of the referenced paragraph without trailing periods; prior levels are omitted unless part of the current level.                                                                          | Word 2003 Help; Microsoft Support. |
| `\p`   | **Above/below indicator.** Shows the position of the REF field relative to the bookmark (“above” or “below”).  Used with `\n`, `\r` or `\w` it appends “above”/“below” to the resulting paragraph number.                                                                     | Word 2003 Help; Microsoft Support. |
| `\r`   | **Relative paragraph number.** Inserts the full paragraph number of the bookmark relative to its position in the numbering scheme, without trailing periods.                                                                                                                  | Word 2003 Help; Microsoft Support. |
| `\t`   | **Suppress non‑numeric text.** When used with `\n`, `\r` or `\w` this switch suppresses any non‑delimiter or non‑numeric text so that only the numeric portion of the paragraph number is displayed (e.g., referencing “Section 1.01” yields “1.01”).                         | Word 2003 Help; Microsoft Support. |
| `\w`   | **Full contextual paragraph number.** Inserts the paragraph number in full context from anywhere in the document.  For example, referencing sub‑paragraph “ii” returns “1.a.ii”.                                                                                              | Word 2003 Help; Microsoft Support. |

In addition to the REF‑specific switches above, Word allows general format switches on most fields—including REF—to control capitalization and formatting. The OOXML specification notes that a REF field may include “one of the following general‑formatting switches: \* Caps, \* FirstCap, \* Lower or \* Upper” followed by zero or one field‑specific switch. These switches correspond to Word’s format switch syntax and determine how letters are capitalised. Other general format switches (\* Charformat to apply the first character’s formatting and \* MERGEFORMAT to preserve previous formatting) can also be appended to a REF field. Numeric (\#) and date/time (\@) format switches are allowed if the REF field result is numeric or a date.


### 5.3 ASK field

`ASK` prompts the user to enter a value and assigns the response to a bookmark.  Syntax:

```text
{ ASK Bookmark "Prompt" [ \d "Default" ] [ \o ] }
```

- `\d "Default"` — supplies a default response if the user presses ENTER without typing a value([ASK field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefASK1.htm)).  To specify an empty default, type `""`([ASK field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefASK1.htm)).
- `\o` — used in mail merges; updates the bookmark in the merge document rather than the main document([ASK field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefASK1.htm)).

After prompting, insert a `REF Bookmark` field to display the value.

Word will prompt for an ASK field even if the bookmark is already defined. From the Word 2003 field-code documentation for ASK:

> “Word displays the prompt each time the ASK field is updated. A response remains assigned to the bookmark until you enter a new response. If the ASK field is used in a mail merge main document, the prompt is displayed each time you merge a new data record unless you use the \o switch.”


### 5.4 DOCVARIABLE field

`DOCVARIABLE` displays the value of a document variable created by a macro or automation code.  Syntax:

```text
{ DOCVARIABLE "Name" }
```

It retrieves the string assigned to the named variable; if the variable does not exist, nothing is displayed([DOCVARIABLE field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-docvariable-field-32a81e22-c5c1-4b16-8097-f0de851db67c)).

### 5.5 DOCPROPERTY field

`DOCPROPERTY` inserts the value of a built‑in or custom document property.  Syntax:

```text
{ DOCPROPERTY "PropertyName" [ \* FormatSwitches ] }
```

It displays the content of the specified property.  You can apply format switches, such as `\* Upper`, to change the case([DOCPROPERTY field (Microsoft Support)](https://support.microsoft.com/en-gb/office/field-codes-docproperty-field-bf00526e-18cd-4515-8c8e-39d59094395a)).

### 5.6 MERGEFIELD

`MERGEFIELD` is used in mail merge documents.  Syntax:

```text
{ MERGEFIELD FieldName [ \b "BeforeText" ] [ \f "AfterText" ] [ \m ] [ \v ] }
```

- `\b` — text inserted **before** the field result when the merge field is not blank([MERGEFIELD field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefMERGEFIELD1.htm)).
- `\f` — text inserted **after** the field result when the merge field is not blank([MERGEFIELD field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefMERGEFIELD1.htm)).
- `\m` — maps the merge field to a pre‑defined data field([MERGEFIELD field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefMERGEFIELD1.htm)).
- `\v` — displays the result vertically (useful for Asian languages)([MERGEFIELD field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefMERGEFIELD1.htm)).

Example: `{ MERGEFIELD FirstName \b " " \f ", " }` inserts a space before and a comma after the first name when it exists.

### 5.7 Sequence (`SEQ`) field

`SEQ` inserts a sequence number that automatically increments each time the field is encountered.  Syntax:

```text
{ SEQ Identifier [ Bookmark ] [ \Switches ] }
```

| Switch | Meaning                                       | Notes                                                                                                                                                                                   |
| ------ | --------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `\c`   | **Repeat closest preceding sequence number.** | Repeats the most recent number for this Identifier (useful in headers/footers for “Chapter X – Page Y”).                                                                                |
| `\h`   | **Hide field result.**                        | Hides the SEQ’s visible result; typically used when you want to cross-reference the sequence without printing the number. Does *not* hide if a general format switch (`\*`) is present. |
| `\n`   | **Next sequence number (default).**           | Inserts the next number in the sequence for this Identifier. If you omit a switch, this is what SEQ does.                                                                               |
| `\r n` | **Reset sequence to `n`.**                    | Resets the sequence to the specified number `n`. Example: `{ SEQ Figure \r 3 }` starts figure numbering at 3.                                                                           |
| `\s n` | **Restart per heading level.**                | Resets the sequence at the outline/heading level given by `n`. For example `{ SEQ Figure \s 2 }` restarts figure numbering at each **Heading 2** section.                               |

Example: to number figures in a report: `{ SEQ Figure \* Arabic }` on each figure caption yields Figure 1, Figure 2, etc.

On top of the SEQ-specific switches, you can attach general formatting switches to the field result:

- Capitalization and number formats via `\*`: 
  - `\* Caps`, `\* FirstCap`, `\* Lower`, `\* Upper` for capitalization
  - `\* alphabetic`, `\* Arabic`, `\* CardText`, `\* DollarText`, `\* Hex`, `\* OrdText`, `\* Ordinal`, `\* roman` for numeric representation (letters, words, roman numerals, etc.)
  - `\* Charformat`, `\* MERGEFORMAT` for character formatting behaviour

- Numeric format switch `\#` for custom numeric pictures (e.g. `\# "000"` or `\# "$#,##0.00"`), if your sequence value is numeric.

---

## 6 Mail Merge‑Specific Conditional Fields

Word provides additional conditional fields for mail merges beyond `IF`:

- **SKIPIF** (section 4.3) skips the current record if the condition is true, thereby not producing a document for that record([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).
- **NEXTIF** moves to the next data record if the condition is true (rarely used).

These fields use the same syntax and operators as the `IF` field and respect the same wildcard rules([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).  `SKIPIF` should not be used with `NEXT` fields([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).

---

## 7 Implementation Considerations for Offline Evaluation

### 7.1 Parsing strategy

A production parser should:

1. **Tokenise** the field code string: identify field boundaries (`{…}`), field names, nested field delimiters, quoted strings and escape sequences.
2. **Build an AST** where each node represents a field type with its parameters and switches.
3. **Maintain a scope for variables/bookmarks**.  `SET` defines variables; `REF` retrieves them; `ASK` prompts user input.  Document variables (from `DOCVARIABLE`) are external and may require an API to query macros.
4. **Evaluate nested fields from the innermost outwards**, substituting results into their parent fields.
5. **Implement comparison semantics** consistent with Word: treat operands as numbers if both are numeric; otherwise compare as strings; respect wildcards; treat absence of FalseText as empty string.
6. **Apply formatting** after computing the raw result: process `\*`, `\#`, and `\@` switches; apply character formatting (case conversion) before numeric/date formatting; apply `\* MERGEFORMAT` to preserve existing formatting.
7. **Handle regional settings** for decimal and digit grouping; allow overrides via configuration.
8. **Respect update triggers**: fields update when the document is opened, printed or when `F9` is invoked.  In offline evaluation, always recompute results when a source variable changes.

### 7.2 Caveats

- **No formal error reporting:** Word silently displays field codes when it cannot evaluate them.  A parser should surface meaningful error messages (e.g., malformed syntax, undefined variables, mismatched braces).
- **Fragile quoting:** unmatched quotes or extra spaces can cause fields to behave unexpectedly.  Always strip outer quotes only when appropriate.
- **Compatibility:** Older `.doc` files and compatibility modes may alter behaviour (e.g., evaluation of numeric vs string).  Test against representative samples.
- **Mail merge data:** In mail merges, `MERGEFIELD` and conditional fields reference external data sources; offline evaluation must provide a record set to simulate merging.

---

## 8 Examples and Templates

### 8.1 Complex conditional formatting

```text
{ IF { MERGEFIELD Amount } > 0
  "{ MERGEFIELD Amount \# "$#,##0.00" } is due"
  "No balance due" } \* Caps
```

If `Amount` is positive, the formatted amount is displayed; otherwise `No balance due`.  The outer `\* Caps` capitalises the first letter of each word.

### 8.2 Nested calculations with variables

```text
{ SET Rate 0.05 }
{ SET Principal { ASK PmtAmt "Enter principal amount" } }
{ SET Interest { = Principal * Rate } }
Total interest due: { REF Interest \# "$#,##0.00" }
```

This snippet prompts the user for a principal amount, calculates interest at 5% and displays it with currency formatting.

### 8.3 Date and time formatting

```text
Document created on { CREATEDATE \@ "dddd, MMMM d, yyyy 'at' h:mm am/pm" }
```

The `CREATEDATE` field is formatted with the full weekday, full month, day and four‑digit year, followed by the time and `AM/PM` notation([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).

---

## References

- **Microsoft Support – Format field results:** Official documentation on the three switch types, capitalisation formats, number formats, character formats, numeric pictures and date/time formats([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c))([Format field results (Microsoft Support)](https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)).
- **Microsoft Support – Set field:** Describes the SET field syntax, bookmark assignment and examples([SET field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-set-field-1fdfbcf9-4d7b-41e2-a1cb-4384a1f516e6)).
- **Microsoft Support – SkipIf field:** Provides syntax and behaviour for SKIPIF, including operator rules and example([SKIPIF field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-skipif-field-d3ff3970-31f3-43a3-be7f-f5fa1704a512)).
- **Microsoft Support – SEQ field:** Describes sequence fields and options for numbering([SEQ field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-seq-sequence-field-062a387b-dfc9-4ef8-8235-29ee113d59be)).
- **Microsoft Support – REF field:** Lists REF field switches and usage([REF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefBOOKMARK1.htm)).
- **Microsoft Support – DOCVARIABLE field:** Explains retrieving document variables([DOCVARIABLE field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-docvariable-field-32a81e22-c5c1-4b16-8097-f0de851db67c)).
- **Microsoft Support – DOCPROPERTY field:** Shows how to insert document properties with format switches([DOCPROPERTY field (Microsoft Support)](https://support.microsoft.com/en-gb/office/field-codes-docproperty-field-bf00526e-18cd-4515-8c8e-39d59094395a)).
- **Microsoft Support – MERGEFIELD:** Defines mail‑merge field switches and examples([MERGEFIELD field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefMERGEFIELD1.htm)).
- **Microsoft Support – ASK field:** Describes ASK field syntax, default values and mail merge behaviour([ASK field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefASK1.htm)).
- **Microsoft Support – Compare field:** Documents the COMPARE field syntax and semantics([COMPARE field (Microsoft Support)](https://support.microsoft.com/en-us/office/field-codes-compare-field-60bfb300-c58d-4f2f-8255-f1a9707390c8)).
- **Microsoft Office Word 2003 Documentation – IF field:** Provides formal syntax, quoting rules, operators and wildcards for IF fields([IF field (Word 2003 documentation)](https://documentation.help/MS-Office-Word-2003/worefIF1.htm)).
- **Microsoft Office Word 2003 Documentation – Examples of IF fields:** Shows complex IF examples combining merge fields, calculations and nested conditions([Examples of IF fields (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefExamplesOfIFFields1.htm)).
- **Microsoft Office Word 2003 Documentation – Formula (=) field:** Describes the formula field operators, functions, cell references and numeric picture switch([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm))([Formula (=) field (Word 2003 documentation)](https://documentation.help/ms-office-word-2003/worefFormula1.htm)).
- **Graham Mayor – Formatting Word fields:** Provides best practices for inserting and updating fields (use Ctrl + F9, Alt + F9, F9) and explains fragility of field codes([Formatting Word fields (Graham Mayor)](https://www.gmayor.com/formatting_word_fields.htm)).

# Appendix – Additional Word Field Codes

## NEXT Field (Mail Merge Record Advance)

The **NEXT** field is a mail-merge control field that forces Microsoft Word to advance the merge data source to the next record unconditionally and merge it into the current output without starting a new section or document. Its syntax is `{ NEXT }`.

Unlike conditional merge fields such as `NEXTIF`, `NEXT` accepts no expressions or switches. Its sole effect is to reposition the merge data source pointer. Word generates this field when the user inserts a *Next Record* rule in the mail-merge UI. It is commonly used in label layouts or multiple-record-per-page designs. The field must appear in the main document body; use in headers, footers or nested contexts may cause incorrect evaluation or record skipping.

**Reference:**  
Microsoft Support – *List of field codes in Word*.  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

---

## MERGEREC and MERGESEQ

Word provides merge-specific counter fields:

- **MERGEREC** inserts the current record number from the underlying data source.
- **MERGESEQ** inserts a sequential counter of records actually merged into the output document, starting at 1 for each merge run.

These fields are distinct from the general `SEQ` field and are typically used to number labels, apply conditional formatting, or display merge order independent of source indexing.

**Reference:**  
Microsoft Support – *List of field codes in Word*.  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

---

## Document Metrics Fields

Word exposes document statistics via the following fields:

- **NUMPAGES** – total number of pages  
- **NUMWORDS** – total word count  
- **NUMCHARS** – total character count  

These fields are frequently used in headers, footers and front matter and are updated when fields are refreshed, printed or previewed.

**Reference:**  
Microsoft Support – *List of field codes in Word*.  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

---

## Cross-Reference Fields

Two common reference-oriented fields are:

- **PAGEREF** – inserts the page number on which a bookmarked item appears.
- **NOTEREF** – inserts the reference mark for a bookmarked footnote or endnote.

These differ from the general `REF` field in that they resolve pagination context (`PAGEREF`) or note markers (`NOTEREF`) rather than arbitrary bookmarked text.

**Reference:**  
Microsoft Support – *List of field codes in Word*.  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

---

## Other Notable Field Types (Overview)

Microsoft Word supports many additional field codes not covered elsewhere in this specification, including:

- Mail-merge helpers such as `AddressBlock` and `GreetingLine`.
- External content fields such as `IncludeText`, `IncludePicture`, `Database` and `Embed`.
- Structural fields such as `TOC`, `Index`, `Citation` and `Bibliography`.
- Automatic numbering and text fields (`AutoNum*`, `AutoText*`).
- Document property fields (`Author`, `Title`, `Keywords`, etc.).

A complete parser or interpreter should be prepared to encounter these fields and either evaluate them or surface appropriate placeholders or errors depending on context.

**Reference:**  
Microsoft Support – *List of field codes in Word*.  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

## GREETINGLINE Field (Implementation-Oriented Description)

The **`GREETINGLINE`** field is a mail-merge convenience field that evaluates to a formatted salutation by combining multiple merge values with fixed text according to a stored greeting pattern. Unlike `MERGEFIELD`, it does not correspond to a single data-source column. Instead, it operates on **role-based merge values**—commonly a courtesy title (for example, “Mr.”, “Mme”), a first name, and a last name—which are mapped by the user at insertion time through Word’s *Mailings → Greeting Line* dialog. These roles are resolved at evaluation time by reading the associated merge fields from the current data record; Word does not infer gender or derive titles and uses the literal values present in the data source.

From an implementation standpoint, `GREETINGLINE` can be treated as a **macro-style field** whose field code encapsulates the selected greeting format, optional fallback text, and language context chosen in the UI. Evaluation consists of selecting the stored greeting pattern (for example, salutation + title + last name, or salutation + first name), substituting the available role values, and inserting the resulting text into the document. If one or more role values are missing or empty, the behavior follows the stored fallback configuration: either emitting a generic greeting string if one was defined, or rendering the greeting with the missing components omitted while preserving surrounding punctuation. Localization affects only the fixed salutation text and punctuation conventions; the field name and role-resolution mechanism are language-independent. For offline or non-interactive processing, an interpreter may model `GREETINGLINE` as syntactic sugar that expands into a sequence of literal text and `MERGEFIELD` evaluations driven by the persisted greeting pattern.

**References:**  
Microsoft Support – *Insert mail merge fields (including Address Block and Greeting Line)*.  
https://support.microsoft.com/en-us/office/insert-mail-merge-fields-9a1ab5e3-2d7a-420d-8d7e-7cc26f26acff  

Microsoft Support – *List of field codes in Word* (alphabetical index, includes `GreetingLine`).  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

## ADDRESSBLOCK Field (Implementation-Oriented Description)

The **`ADDRESSBLOCK`** field is a mail-merge convenience field that evaluates to a formatted postal address block by aggregating multiple address-related merge values into a single, multi-line output. Unlike `MERGEFIELD`, it does not correspond to a single data-source column. Instead, it operates on a set of **logical address roles**—such as recipient name, company, street address, city, state/province, postal code, and country—that are mapped by the user at insertion time through Word’s *Mailings → Address Block* dialog and, if necessary, the associated *Match Fields* interface.

From an implementation perspective, `ADDRESSBLOCK` can be treated as a **macro-style composite field** whose evaluation consists of expanding a stored address template over the mapped address roles. The template defines the ordering of address components, the insertion of line breaks, and the placement of punctuation (for example, commas or spaces between locality elements). During evaluation, the field retrieves the mapped merge values for the current data record and substitutes them into the template. Components whose mapped values are empty are omitted from the output, typically along with their associated separators, resulting in a compact address block without blank lines.

The field does not expose user-editable switches in the field code for controlling layout; formatting choices are captured implicitly from the insertion dialog rather than through explicit `\\switch` syntax. Localization affects the default ordering of address components and country-specific formatting conventions but does not alter the field name or its fundamental role-resolution mechanism. For offline or non-interactive processing, an interpreter may model `ADDRESSBLOCK` as syntactic sugar that expands into a sequence of `MERGEFIELD` evaluations and literal line breaks according to the persisted address template and role mappings.

**References:**  
Microsoft Support – *Insert mail merge fields (including Address Block)*.  
https://support.microsoft.com/en-us/office/insert-mail-merge-fields-9a1ab5e3-2d7a-420d-8d7e-7cc26f26acff  

Microsoft Support – *How to use the Mail Merge feature in Word* (includes Address Block and field matching).  
https://support.microsoft.com/en-us/topic/how-to-use-the-mail-merge-feature-in-word-to-create-and-to-print-form-letters-that-use-the-data-from-an-excel-worksheet-d8709e29-c106-2348-7e38-13eecc338679  

Microsoft Support – *List of field codes in Word* (alphabetical index, includes `AddressBlock`).  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51


## DATABASE Field (Implementation-Oriented Description)

The **`DATABASE`** field is a Word field code that allows a document to retrieve data dynamically by executing a query against an external database and inserting the resulting dataset into the document, typically as a table. It is part of Word’s standard field code set and is most commonly used in advanced mail-merge or reporting scenarios where data must be selected or filtered independently of the primary merge data source.

Functionally, the field encapsulates two core elements: a **database connection specification** and a **query string**. The connection information is provided via a field switch (commonly `\\d`) and may reference a local file (such as an Access `.mdb`/`.accdb` or Excel workbook) or a full ODBC/OLE DB connection string. When required by the target database engine, this connection string may include authentication parameters such as user name and password or may rely on integrated security. Word does not interpret or secure these credentials; it passes the connection string verbatim to the underlying database provider.

The query itself is supplied as a literal string (commonly via the `\\s` switch) and is executed by the database engine specified in the connection, not by Word. As a result, the query language is **not Word-specific SQL** but rather the SQL dialect supported by the target system (for example, Access SQL for Jet/ACE databases, Transact-SQL for Microsoft SQL Server, or another vendor-specific dialect when using ODBC). Standard SQL constructs such as `SELECT`, `FROM`, and `WHERE` are generally supported, while advanced features depend entirely on the database backend.

A notable characteristic of the `DATABASE` field is that the query string may contain **embedded Word fields**, such as `MERGEFIELD`, allowing the query to be parameterized based on the current merge record or document state. In this case, embedded fields are evaluated first, and their results are substituted into the query string before the database query is executed. The resulting recordset is then rendered into the document, usually as a table whose rows and columns correspond to the query output. Formatting and layout options are limited and largely controlled by Word rather than by the field code itself.

For offline or programmatic evaluation, `DATABASE` should be treated as an **external-data macro field**. A compatible implementation involves resolving embedded fields, establishing a database connection using the supplied connection string, executing the query using the appropriate database driver, and inserting the returned rows into the document model. Because the field depends on external systems and credentials, implementations commonly choose to support it conditionally, provide a stubbed evaluation, or surface a controlled error when database access is unavailable.

**References:**  
Microsoft Support – *List of field codes in Word* (alphabetical index, includes `Database`).  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51  

BetterSolutions – *Database field* (overview and common switches).  
https://bettersolutions.com/word/fields/database-field.htm  

Stack Overflow – *Using Word DATABASE field with SQL and embedded MERGEFIELD*.  
https://stackoverflow.com/questions/73159646/ms-word-database-field-sql-switch-how-to-reference-merge-field

## Field Error Semantics (Implementation-Oriented Description)

Microsoft Word reports field-evaluation failures by substituting the expected field result with a **literal, human-readable error string**. These error strings are inserted directly into the document and thereafter behave as ordinary text: they participate in nested field evaluation, conditional logic, comparisons, and variable assignment. Word does not expose numeric error codes or structured exceptions at the field level; error reporting is entirely text-based and user-facing.

When a field fails to evaluate (for example, due to a missing bookmark, an unresolved cross-reference, or an invalid structural context), Word replaces the field result with an error message beginning with a localized equivalent of the prefix **“Error!”**. Evaluation does not halt. Instead, the error text becomes the field’s result and is propagated outward if the field is nested inside another field. As a consequence, conditional fields such as `IF`, `COMPARE`, or formulas operate on the literal error text rather than on a null or undefined value.

Commonly observed English-language field error messages include:

- **“Error! Bookmark Not Defined.”** — a referenced bookmark does not exist (e.g., `REF` to a deleted bookmark).  
- **“Error! Reference source not found.”** — a cross-reference target (heading, caption, or numbered item) cannot be resolved.  
- **“Error! No text of specified style in document.”** — a field depending on the presence of a specific style (for example, certain numbering or caption contexts) cannot find matching content.  

Microsoft does not publish an exhaustive list of all possible field error strings. In practice, most field-related errors share a common pattern: a localized form of **“Error!”** followed by a short diagnostic phrase. Because these messages are localized, their exact wording depends on the Word UI language. For example, in French builds of Word, “Error! Bookmark Not Defined.” appears as **“Erreur ! Signet non défini.”**. Implementations must therefore avoid hard-coding English error strings and instead rely on pattern-based or locale-aware detection.

From an implementation perspective, field errors should be modeled with the following semantics:
- A field error yields a **string result**, not an exception.
- The error string is **propagated unchanged** into any enclosing field.
- Comparisons and calculations that consume an error result treat it as text; numeric coercion does not occur.
- Error detection should be **locale-aware**, typically by matching a configurable set of localized “Error!” prefixes or known message patterns.

For offline or programmatic evaluation, it is advisable to distinguish internally between a *successful evaluation result* and an *error-result string*, even though both ultimately render as text. This allows downstream consumers to detect error conditions while preserving Word-compatible behavior in nested evaluation and conditional logic.

**References:**  
Microsoft Support – *Troubleshoot bookmarks* (documents “Error! Bookmark Not Defined.”).  
https://support.microsoft.com/en-us/office/troubleshoot-bookmarks-9cad566f-913d-49c6-8d37-c21e0e8d6db0  

Microsoft Learn / Support forums – *Reference source not found* (cross-reference errors).  
https://learn.microsoft.com/en-us/answers/questions/5638855/in-a-shared-word-document-how-to-fix-not-being-abl  

Super User – *Error! No text of specified style in document* (community-documented field error).  
https://superuser.com/questions/641315/ms-word-field-code-error-error-no-text-of-specified-style-in-document


## External Content Fields: INCLUDETEXT and INCLUDEPICTURE

### Overview

The `INCLUDETEXT` and `INCLUDEPICTURE` fields are external-content fields that dynamically incorporate content from outside the current Word document at field-evaluation time. Unlike mail-merge fields, which bind to a structured data source, these fields reference external files or resources and inject their contents directly into the document model. They are explicitly listed in Microsoft’s canonical field index and behave as macro-style fields whose evaluation depends on filesystem or URI access rather than merge records.

### INCLUDETEXT Field

The `INCLUDETEXT` field inserts text from another document. Its general syntax is:

    { INCLUDETEXT "FullFilePath" [BookmarkName] [\\!] }

- **FullFilePath** is a quoted absolute or relative path to a Word document (`.doc`, `.docx`, `.rtf`). Paths containing spaces must be quoted.
- **BookmarkName** is optional. If present, only the content delimited by the named bookmark in the source document is included. If omitted, the entire document body is inserted.
- **\\! (lock nested fields)** suppresses automatic updating of fields contained within the included text. Without this switch, Word evaluates nested fields in the included content when the `INCLUDETEXT` field is updated.

Evaluation semantics:
1. Word resolves the file path and opens the external document.
2. If a bookmark is specified, Word extracts only that bookmarked range; otherwise, it extracts the full document body.
3. The extracted content is inserted verbatim into the current document.
4. Nested fields inside the included content are evaluated unless the `\\!` switch is present.

The included content is not copied permanently; it is refreshed whenever fields are updated. Missing files or bookmarks result in literal error text being displayed at the field location.

### INCLUDEPICTURE Field

The `INCLUDEPICTURE` field inserts a picture from an external file or URI. Its general syntax is:

    { INCLUDEPICTURE "URIorFilePath" [\\d] }

- **URIorFilePath** may be a local filesystem path or a remote URL.
- **\\d (link instead of embed)** instructs Word to store only a link to the image rather than embedding the image binary data in the document.

Supported formats include common raster image types such as PNG, JPG/JPEG, GIF, BMP, and TIFF. The field does not extract images from PDF files; PDFs must be converted to supported image formats before inclusion.

Evaluation semantics:
1. Word resolves the path or URI.
2. The image is loaded at field-update time.
3. The image is rendered as a picture object in the document.
4. If `\\d` is present, the document retains only a reference to the external image and reloads it on subsequent updates.

### Formatting and Layout

Neither `INCLUDETEXT` nor `INCLUDEPICTURE` exposes field-code switches for controlling layout, sizing, alignment, borders, cropping, or positioning.

- For `INCLUDETEXT`, formatting is inherited from the source document and from the surrounding paragraph context in the target document.
- For `INCLUDEPICTURE`, all visual formatting (size, scale, wrapping, borders, effects) is applied through Word’s picture object properties after the field has resolved.

From an implementation perspective, formatting is not part of field evaluation and must be modeled as document-object properties associated with the inserted content.

### Implementation Notes

For offline or programmatic evaluation:
- Treat both fields as external-resource resolvers.
- Resolve nested fields before or after inclusion according to the `\\!` switch.
- Guard against recursive inclusion loops.
- Handle missing files, invalid URIs, and inaccessible resources by emitting Word-style error text rather than throwing hard errors.
- Consider security implications when resolving external paths or URLs.

### References

Microsoft Support – List of field codes in Word  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

Microsoft Office Word 2003 Documentation – INCLUDETEXT field  
https://documentation.help/ms-office-word-2003/worefINCLUDETEXT1.htm

BetterSolutions – INCLUDETEXT and INCLUDEPICTURE field behavior  
https://bettersolutions.com/word/fields/includetext-field.htm  
https://bettersolutions.com/word/fields/includepicture-field.htm


## Structural and Aggregate Fields: TOC, TC, INDEX, XE, RD

### Overview

The TOC/INDEX ecosystem in Microsoft Word consists of *aggregate fields* that synthesize document-wide structures by scanning, collecting, sorting, and rendering information distributed across the document (and optionally across multiple documents). These fields differ fundamentally from scalar fields (e.g., REF, DATE) in that their evaluation requires a **multi-pass document scan** and access to pagination metadata.

The ecosystem is composed of:

- **TOC** — Table of Contents generator
- **TC** — Table of Contents entry marker
- **INDEX** — Index generator
- **XE** — Index entry marker
- **RD** — External document reference for aggregation

Evaluation of TOC and INDEX fields depends on the presence, placement, and configuration of TC/XE/RD fields and paragraph metadata.

---

## TOC Field (Table of Contents)

### Syntax (canonical form)

    { TOC [\\o "start-end"] [\\t "Style,Level;…"] [\\u] [\\b Bookmark] [\\f Identifier] [\\l Levels] [\\c SeqId] [\\a CaptionId] }

### Data sources scanned

A TOC field aggregates entries from one or more of the following sources, depending on switches:

1. **Paragraph outline levels** (typically via built-in Heading styles).
2. **Paragraph outline attributes** (explicit outline level property, via \\u).
3. **TC fields** whose identifiers and levels match TOC filters.
4. **Captioned SEQ fields** when \\c or \\a is used.
5. **Referenced documents** via RD fields.

### Entry model

Each TOC entry is resolved to:

- **Display text**
- **Hierarchy level** (integer ≥ 1)
- **Target location** (paragraph anchor)
- **Resolved page number**

Entries are ordered primarily by document order, not alphabetically.

### Switch semantics

- **\\o "n-m"** — Include headings with outline levels between *n* and *m*.
- **\\t "Style,Level;…"** — Map paragraph styles to explicit TOC levels.
- **\\u** — Include paragraphs with an outline level property even if not styled.
- **\\b Bookmark** — Restrict scanning to the bookmarked document range.
- **\\f Identifier** — Include only TC fields tagged with this identifier.
- **\\l Levels** — Include only TC entries whose \\l value matches Levels.
- **\\c SeqId** — Build TOC from SEQ fields with the given identifier.
- **\\a CaptionId** — Include captioned objects of a given type (text only).

### Evaluation notes

- Pagination must be resolved before final TOC rendering.
- TOC output is regenerated wholesale on update.
- Layout (leaders, indentation) is controlled by styles, not field switches.

---

## TC Field (Table of Contents Entry Marker)

### Syntax

    { TC "EntryText" \\l Level [\\f Identifier] }

### Semantics

- TC fields are **non-rendering markers**.
- They inject synthetic TOC entries independent of paragraph styles.
- EntryText is literal and not derived from document content.
- Level defines TOC nesting depth.
- Identifier allows partitioning multiple TOCs.

During TOC evaluation, TC fields are collected, filtered, and converted into TOC entries as if they were headings at the specified level.

---

## INDEX Field

### Syntax (simplified)

    { INDEX [\\b Bookmark] [\\f Identifier] [\\e Separator] [\\h] [\\r] [\\z] }

### Data sources scanned

- All XE fields within scope.
- XE fields from referenced documents via RD.

### Entry model

Each index entry resolves to:

- **Primary term**
- Optional **secondary term**
- Optional **cross-reference text**
- One or more **page numbers**

Entries are **alphabetically sorted** using locale-dependent collation.

### Evaluation semantics

- Duplicate page references for the same term are collapsed.
- Cross-references are rendered as “See” / “See also” constructs.
- Page numbers are resolved post-layout.

---

## XE Field (Index Entry Marker)

### Syntax

    { XE "Term[:Subterm]" [\\t "CrossRefText"] [\\f Identifier] }

### Semantics

- XE fields mark indexable terms at a document location.
- Primary and secondary terms are parsed from the string.
- Multiple XE fields with the same term accumulate page numbers.
- Cross-reference entries suppress page numbers and emit reference text.

XE fields do not render content directly.

---

## RD Field (Referenced Document)

### Syntax

    { RD "DocumentPath" }

### Semantics

- RD fields extend the scan scope of TOC and INDEX.
- Referenced documents are opened and scanned as read-only sources.
- TC/XE/heading metadata is merged into the parent aggregation.
- Pagination is resolved in the context of the *source document*, not the host.

Recursive RD chains must be cycle-checked by the implementation.

---

## Processing Model (Implementation Guidance)

A Word-compatible implementation requires:

1. **Pre-scan phase** — identify TC, XE, RD, heading metadata.
2. **Recursive document loading** — resolve RD references.
3. **Entry collection** — normalize entries into TOC/INDEX models.
4. **Filtering** — apply identifier, bookmark, and level constraints.
5. **Sorting** — structural (TOC) or locale-aware collation (INDEX).
6. **Pagination resolution** — requires layout engine integration.
7. **Rendering** — generate structured output using styles.

These fields cannot be evaluated correctly without document-level context and pagination awareness.

---

## References

Microsoft Support – List of field codes in Word  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

Microsoft Support – Create and update a table of contents  
https://support.microsoft.com/en-us/office/create-a-table-of-contents-or-update-a-table-of-contents-882e8564-0edb-435e-84b5-1d8552ccf0c0

Microsoft Support – Create and update an index  
https://support.microsoft.com/en-us/office/create-and-update-an-index-cc502c71-a605-41fd-9a02-cda9d14bf073

Office-Watch – All Word TOC switches explained  
https://office-watch.com/2023/all-word-table-of-contents-options/


## Header, Footer, and Section Scope Semantics

In Microsoft Word, headers and footers are not part of the main document text flow but are stored as **section-scoped story ranges** with independent evaluation contexts. Each section maintains up to six distinct header/footer containers (primary, first-page, and even-page variants for both header and footer). Fields located in these containers are evaluated against the **section’s pagination and numbering state**, not the linear body text stream. As a result, the same field code (for example `{ PAGE }`) can yield different results in the body versus in a header or footer, depending on the section’s page-numbering configuration, restart rules, and linkage settings.

Section breaks establish hard evaluation boundaries. Fields such as `PAGE`, `SECTION`, `SECTIONPAGES`, `NUMPAGES`, and `SEQ \\s n` resolve their values relative to the **current section context** when placed in headers or footers. For example, `{ PAGE }` reflects the page index within the section’s numbering scheme (including any “Start at” offset), while `{ NUMPAGES }` returns the total page count for the entire document, and `{ SECTIONPAGES }` returns the total page count of the current section only. When headers or footers are marked as *linked to previous*, their field content and evaluation context are effectively inherited from the preceding section; when unlinked, the fields are evaluated independently using the new section’s metadata.

Field update behavior also differs structurally. Bulk updates such as “select all + F9” operate on a single story range and therefore do not reliably update fields in headers, footers, text boxes, or other non-body stories. In Word automation and in practice, headers and footers require an explicit traversal and update pass per section. Mail-merge control fields (`NEXT`, `NEXTIF`, `SKIPIF`) are not well-defined in headers or footers and are only consistently supported in the main document body; their placement outside the body yields version-dependent or undefined behavior.

For an implementation that aims to match Word, headers and footers must be modeled as **separate field containers bound to section objects**, with explicit linkage state, independent field lists, and access to section-level pagination data. Evaluation of header/footer fields must be deferred until pagination is known, and section boundaries must invalidate and recompute dependent field results when numbering or layout changes occur.

### References

Microsoft Support – List of field codes in Word  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51

AddBalance – Word sections and headers/footers (technical overview)  
https://www.addbalance.com/usersguide/sections.htm

MakeOfficeWork – Page numbering and section behavior  
https://makeofficework.com/showing_page_numbers.htm

Microsoft public Word VBA discussions – Updating fields in headers/footers  
https://microsoft.public.word.vba.general.narkive.com/HoWXuFiP/updating-fields-in-headers-footers


## Unicode, RTL, and Vertical Text Semantics in Field Evaluation

Microsoft Word evaluates field codes over **logical Unicode text** (internally UTF-16) and applies script directionality, vertical layout, and glyph shaping strictly at the **rendering stage**, not during field evaluation. Field results are always produced as linear Unicode strings; bidirectional reordering, character mirroring, and vertical glyph orientation are applied later based on paragraph, run, and page layout properties. This separation is critical: fields do not alter character order for right-to-left (RTL) scripts, nor do they insert layout control characters as part of evaluation.

In RTL contexts (for example Arabic or Hebrew paragraphs), fields inherit the paragraph’s directional property. Mixed LTR/RTL content within a field result is rendered according to the Unicode Bidirectional Algorithm, with numbers and Latin text maintaining left-to-right ordering unless explicitly overridden by embedding controls. This visual reordering does not affect the underlying string value returned by the field and therefore must be treated as a layout concern by any implementation. Nested fields, comparisons, and numeric formatting operate on the logical string and numeric values, independent of visual direction.

Case-conversion format switches (`\\* Upper`, `\\* Lower`, `\\* FirstCap`, `\\* Caps`) operate on Unicode characters and are **locale-sensitive** where applicable. While many scripts are unicameral and unaffected (for example Arabic or Hebrew), Latin script segments embedded in RTL text are subject to language-specific case mappings. An implementation that aims to match Word behavior should therefore rely on Unicode case-mapping tables with locale awareness rather than ASCII-only transformations.

Vertical text handling is primarily relevant to East Asian layouts (Japanese, Chinese, Korean) and is orthogonal to field semantics. The `\\v` switch on `MERGEFIELD` and related fields does not modify the semantic value of the field; instead, it acts as a **rendering hint** indicating that the result should participate correctly in vertical text flow when the document or paragraph is set to vertical writing mode. In such layouts, Word may rotate glyphs, select alternate glyph forms, or adjust spacing, but the evaluated field result remains a standard Unicode string. In horizontal layouts, the `\\v` switch typically has no observable effect.

Numeric (`\\#`) and date/time (`\\@`) format switches are not script-specific but are **locale-dependent**. Digits, decimal separators, digit grouping symbols, calendar systems, and month/day names are resolved using the document or system locale. In RTL documents, numerals are often rendered left-to-right within an RTL run, following Unicode bidi rules, without altering their numeric value or string representation.

For implementers, the correct model is therefore layered: field evaluation produces a Unicode string or numeric value; locale-aware formatting is applied next; and bidirectional reordering or vertical layout adjustments are handled exclusively by the rendering engine. Any attempt to combine evaluation and layout logic risks diverging from Word’s behavior, particularly in mixed-script or vertical-text documents.

**References:**  
Microsoft Support – Format field results (Unicode, locale, and formatting behavior)  
https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c  

Superuser – Purpose of the `\\v` switch in MERGEFIELD (vertical text context)  
https://superuser.com/questions/1212759/whats-the-purpose-of-mergefield-v-switch-in-ms-word  

Unicode Consortium – Unicode Bidirectional Algorithm (UAX #9)  
https://www.unicode.org/reports/tr9/

## AUTONUM / AUTONUMOUT / AUTONUMLGL (Detailed Legacy Semantics)

The **AUTONUM**, **AUTONUMOUT**, and **AUTONUMLGL** fields are legacy automatic numbering fields whose values are generated by Word’s internal layout and outline engine rather than by explicit field parameters. Unlike `SEQ` or list styles, these fields expose no switches to control formatting, restarting, or scoping. Their behavior is driven implicitly by **paragraph position, outline level, and document structure**, and updates occur as part of Word’s pagination and layout recalculation rather than through explicit field updates.

`AUTONUM` produces a flat, sequential number that increments in document order, historically used for margin numbering or simple automatic numbering in early Word templates. `AUTONUMOUT` and `AUTONUMLGL` both participate in **outline-level numbering**, deriving their hierarchical structure from the paragraph’s outline level (for example Heading 1, Heading 2, etc.). The distinction between them is representational rather than structural: `AUTONUMOUT` emits outline numbers with trailing punctuation (such as `1.`, `1.1.`, `1.1.1.`), whereas `AUTONUMLGL` emits legal-style hierarchical numbers without trailing punctuation (`1`, `1.1`, `1.1.1`).

These fields update automatically when document structure changes (paragraph insertion, deletion, outline level changes) and are **not reliably refreshed by F9**. Their results may be suppressed, unstable, or absent when nested inside conditional fields (`IF`, `SKIPIF`) or when placed in non-body story ranges. Microsoft does not publish a formal, version-stable specification for AUTONUM* evaluation, and observed behavior varies across Word versions and compatibility modes.

For implementation purposes, the AUTONUM* family should be treated as **non-deterministic, layout-driven generators**. A compatible engine should recognize these field types and either delegate evaluation to Word, approximate numbering using outline metadata, or explicitly flag them as legacy constructs whose exact behavior cannot be reproduced outside Word’s layout engine.

### AUTONUM* Field Comparison (Operational View)

| Field | Increment Driver | Hierarchy Source | Output Form | Distinguishing Characteristic |
|------|------------------|------------------|-------------|-------------------------------|
| AUTONUM | Insertion order | None (flat) | 1, 2, 3 | Simple sequential numbering |
| AUTONUMOUT | Outline level | Paragraph outline level | 1., 1.1., 1.1.1. | Hierarchical numbering with punctuation |
| AUTONUMLGL | Outline level | Paragraph outline level | 1, 1.1, 1.1.1 | Hierarchical numbering without punctuation (legal style) |

**References:**  
Microsoft Support – List of field codes in Word  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51  

Microsoft Learn – Word field type enumeration (WdFieldType)  
https://learn.microsoft.com/en-us/office/vba/api/word.wdfieldtype  

Microsoft Office Word 2003 Documentation – AUTONUMLGL field  
https://documentation.help/MS-Office-Word-2003/worefAUTONUMLGL1.htm

## CITATION and BIBLIOGRAPHY Fields (Implementation-Oriented Semantics)

The **CITATION** and **BIBLIOGRAPHY** fields implement Word’s internal bibliographic system. Unlike most other fields, they do not primarily reference document text, bookmarks, or merge data. Instead, they resolve against a **structured bibliography source store** maintained by Word (the *Source Manager*), where each bibliographic entry is identified by a unique **Tag** and stored as structured metadata (author, title, year, publisher, etc.). Both fields are tightly coupled to Word’s citation-style engine and are not fully specifiable using field code syntax alone.

### CITATION Field

**Syntax:**  
`{ CITATION "Tag" [ switches ] }`

The CITATION field inserts an in-text reference to a single bibliographic source identified by its Tag. The Tag must correspond to an existing source entry in the document’s bibliography source list; if no such source exists, Word renders an error or an empty result depending on version and update context.

At evaluation time, Word performs the following steps:
1. Locate the source record matching the Tag in the active document or attached template.
2. Select the active citation style (APA, MLA, Chicago, ISO 690, etc.).
3. Format the in-text citation according to that style and the current locale.
4. Apply any field-level switches to the formatted result.

**Supported switches (observed and documented):**

| Switch | Meaning |
|------|--------|
| `\\f "text"` | Inserts literal prefix text immediately before the citation. |
| `\\s "text"` | Inserts literal suffix text immediately after the citation. |
| `\\l LCID` | Overrides the locale used to format the citation (language, punctuation, labels). |
| `\\m "Tag2"` | Adds an additional source Tag, allowing multiple sources in a single citation. |
| `\\v number` | Supplies a volume number to the citation, consumed by styles that support it. |

**Example:**  
`{ CITATION "Doe2020" \\l 1033 \\f "see " \\s ", esp. ch. 2" }`

The field code itself does not encode author names, dates, or titles; all such data is retrieved from the source store. Formatting logic (parentheses, author-year ordering, separators) is entirely style-driven.

### BIBLIOGRAPHY Field

**Syntax:**  
`{ BIBLIOGRAPHY [ switches ] }`

The BIBLIOGRAPHY field generates a complete bibliography or works-cited section by aggregating all sources referenced by CITATION fields in the document. It does not take source Tags, as its scope is document-wide.

At evaluation time, Word:
1. Collects all distinct source records referenced by CITATION fields.
2. Sorts and formats them according to the active bibliography style.
3. Emits a formatted block of text representing the bibliography.

**Supported switches:**

| Switch | Meaning |
|------|--------|
| `\\l LCID` | Overrides the locale used for bibliography formatting. |

**Example:**  
`{ BIBLIOGRAPHY \\l 1036 }`  
Formats the bibliography using French locale rules where supported by the selected style.

The BIBLIOGRAPHY field expands to formatted text only; it does not emit nested fields for individual entries. Updating citations or changing the bibliography style requires explicit field refresh to regenerate the output.

### Construction and Storage Model

Both CITATION and BIBLIOGRAPHY rely on:
- A structured source store embedded in the document or template (often serialized as XML).
- A citation-style definition, typically implemented internally or via XSL templates installed with Word.
- Locale metadata that affects labels, punctuation, and ordering.

Because citation formatting logic is external to the field code and driven by Word’s internal style engine, **full reimplementation requires replicating Word’s source schema, style engine, and locale handling**. As a result, these fields are best classified as **engine-dependent macro fields** rather than declarative field expressions.

### Implementation Guidance

For non-Word implementations:
- Parse and preserve CITATION and BIBLIOGRAPHY field codes.
- Model source data as structured records keyed by Tag.
- Apply style-based formatting using a citation processor if full fidelity is required.
- Otherwise, surface these fields as externally-resolved constructs or flag them as non-reproducible without Word’s bibliography engine.

**References:**  
Microsoft Support – List of field codes in Word  
https://support.microsoft.com/en-gb/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51  

Microsoft Support – Add citations and create a bibliography  
https://support.microsoft.com/en-us/office/add-citations-in-a-word-document-ab9322bb-a8d3-47f4-80c8-63c06779f127  

Microsoft Learn – Working with bibliographies in Word  
https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-bibliographies


