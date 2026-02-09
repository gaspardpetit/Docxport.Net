<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 9 -->
<!-- Created: 2026-02-08 19:39:00Z -->
<!-- Modified: 2026-02-08 23:30:00Z -->

Expect 1:

`REF Var1`

1

Expect Error: <b>

`REF VarUnknown`

Error! Reference source not found.</b>

Expect No Error:

`IF  = "" "Empty" "Not Empty"`

Not Empty

Expect one:

`IF  = "1" "one" "not one"`

one

Expect <b>one</b> (bold): <b>

`IF  = "1" "one" "not one"`

one</b>

Expect <b>1</b><u>2</u><b>3:

`IF 1 = 1 "123" "error"`

1</b><u>2</u><b>3</b>