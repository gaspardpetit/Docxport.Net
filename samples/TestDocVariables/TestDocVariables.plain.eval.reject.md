<!-- Title: Field Test Title -->
<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 10 -->
<!-- Created: 2026-02-11 08:11:00Z -->
<!-- Modified: 2026-02-11 09:25:00Z -->
<!-- CustomProp1: custom-value -->
<!-- Created: 2010-10-13T04:00:00Z -->

DOCVARIABLE TEST

<table>
  <tr>
    <td>DOCVARIABLE Var1

</td>
    <td>two

</td>
  </tr>
  <tr>
    <td>DOCVARIABLE Missing

</td>
    <td>Error! No document variable supplied.

</td>
  </tr>
  <tr>
    <td><b>DOCVARIABLE</b> Var1 \* Charformat

</td>
    <td><b>two</b>

</td>
  </tr>
  <tr>
    <td>DOCVARIABLE Var1 \* MERGEFORMAT

</td>
    <td><u>two</u>

</td>
  </tr>
</table>
IF TEST

<table>
  <tr>
    <td>IF 5 >= 3 "OK" "BAD"

</td>
    <td>OK

</td>
  </tr>
  <tr>
    <td>IF "Approved" = "App*" "YES" "NO"

</td>
    <td>YES

</td>
  </tr>
  <tr>
    <td>SET Var2 "two"

</td>
    <td>

</td>
  </tr>
  <tr>
    <td>IF { REF Var1 } = "two" "Value is { REF Var1 }" "Mismatch"

</td>
    <td>Value is two

</td>
  </tr>
  <tr>
    <td>IF { REF MissingBookmark } = "" "Empty" "Not Empty"

</td>
    <td>Not Empty

</td>
  </tr>
</table>
DOCPROPERTY TEST

<table>
  <tr>
    <td>DOCPROPERTY Title

</td>
    <td>Field Test Title

</td>
  </tr>
  <tr>
    <td>DOCPROPERTY Created \@ "yyyy-MM-dd"

</td>
    <td>2010-10-13

</td>
  </tr>
  <tr>
    <td>DOCPROPERTY "CustomProp1"

</td>
    <td>custom-value

</td>
  </tr>
</table>
TEST COMPARE

<table>
  <tr>
    <td>COMPARE 5 >= 3

</td>
    <td>1

</td>
  </tr>
  <tr>
    <td>COMPARE 2 > 3

</td>
    <td>0

</td>
  </tr>
  <tr>
    <td>COMPARE "Approved" = "App*"

</td>
    <td>1

</td>
  </tr>
</table>
TEST MERGEFIELD

<table>
  <tr>
    <td>MERGEFIELD FirstName \b "Hello " \f "!"

</td>
    <td>Hello Ana!

</td>
  </tr>
  <tr>
    <td>MERGEFIELD EmptyField \b "Hello " \f "!"

</td>
    <td>

</td>
  </tr>
  <tr>
    <td>MERGEFIELD GivenName \m

</td>
    <td>Ana

</td>
  </tr>
  <tr>
    <td>MERGEFIELD FirstName \v

</td>
    <td>A
n
a

</td>
  </tr>
  <tr>
    <td>MERGEFIELD FirstName \* Upper

</td>
    <td>ANA

</td>
  </tr>
</table>