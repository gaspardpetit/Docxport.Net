<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 9 -->
<!-- Created: 2026-02-08 19:39:00Z -->
<!-- Modified: 2026-02-08 23:30:00Z -->

Expect 1:

`SET Var1 "1"`

1

`REF Var1`

1

Expect Error:

`REF VarUnknown`

<b>Error! Reference source not found.</b>

Expect No Error:

`IF`



`REF VarUnknow`



`Error! Reference source not found.`



`= "" "Empty" "Not Empty"`

Not Empty

Expect one:

`SET Var1 "1"`

1

`IF`



`REF Var1`



`1`



`= "1" "one" "not one"`

one

Expect <b>one</b> (bold):

`SET Var1 "`



`1`



`"`

1

`IF`



`REF Var1`



`1`



`= "1" "`



`one`



`" "`



`not one`



`"`

<b>one</b>

Expect <b>1</b><u>2</u><b>3:

`IF 1 = 1 "`



`1`



`2`



`3`



`" "error"`

1</b><u>2</u><b>3</b>