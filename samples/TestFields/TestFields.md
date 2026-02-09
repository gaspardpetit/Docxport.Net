<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 9 -->
<!-- Created: 2026-02-08 19:39:00Z -->
<!-- Modified: 2026-02-08 23:30:00Z -->

<div class="section" style="color:#000000;display:flex;flex-direction:column;position:relative;width:8.5in;min-height:11in;box-sizing:border-box;padding-left:1in;padding-right:1in;background-color:#ffffff;font-family:Aptos;font-size:12pt;">
<div class="body" style="flex:1 0 auto;padding-top:1in;">
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

<a id="Var1" data-bookmark-id="0"></a>1

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

</div>
</div>