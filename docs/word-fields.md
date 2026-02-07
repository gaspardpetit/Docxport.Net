# Microsoft Word Field Code Syntax

## Introduction

Microsoft Word’s field code system is an **ad‑hoc mini‑language** for conditional logic, calculations, mail merges and dynamic document content.  The language evolved incrementally; there is **no formal grammar** or unified specification.  Field codes are encapsulated between curly braces (`{ }`) inserted with **Ctrl + F9** and are evaluated by Word when a document is opened, printed or when the user presses **F9** to update fields.  Because the syntax and evaluation rules are poorly documented and fragile, this document consolidates authoritative information from Microsoft documentation and trusted secondary sources to provide an exhaustive reference for implementing a **parser/interpreter** for Word fields.

Field codes consist of:

- A **field type** (`IF`, `SET`, `REF`, `SEQ`, `DATE`, `COMPARE`, etc.).
- An optional **expression or parameters** whose syntax depends on the field type.
- Optional **switches** beginning with a backslash (`\`) that modify the way results are formatted.

This specification is organised into conceptual sections with examples and references.  Each section lists applicable field types, syntax rules, supported parameters and switch behaviour.  A final section discusses implementation considerations for offline evaluation.

---

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

