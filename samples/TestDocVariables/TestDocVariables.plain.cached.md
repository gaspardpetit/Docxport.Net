<!-- Title: Field Test Title -->
<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 11 -->
<!-- Created: 2026-02-11 08:11:00Z -->
<!-- Modified: 2026-02-11 09:33:00Z -->
<!-- CustomProp1: custom-value -->
<!-- Created: 2010-10-13T04:00:00Z -->

DOCVARIABLE TEST

<table>
  <tr>
    <td>DOCVARIABLE Var1

</td>
    <td>one

</td>
  </tr>
  <tr>
    <td>DOCVARIABLE Missing

</td>
    <td><b>Error! No document variable supplied.</b>

</td>
  </tr>
  <tr>
    <td><b>DOCVARIABLE</b> Var1 \* Charformat

</td>
    <td><b>one</b>

</td>
  </tr>
  <tr>
    <td>DOCVARIABLE Var1 \* MERGEFORMAT

</td>
    <td><u>one</u>

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
    <td>REF

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
    <td>Hello «FirstName»!

</td>
  </tr>
  <tr>
    <td>MERGEFIELD EmptyField \b "Hello " \f "!"

</td>
    <td>Hello «EmptyField»!

</td>
  </tr>
  <tr>
    <td>MERGEFIELD GivenName \m

</td>
    <td>«GivenName»

</td>
  </tr>
  <tr>
    <td>MERGEFIELD FirstName \v

</td>
    <td>«FirstName»

</td>
  </tr>
  <tr>
    <td>MERGEFIELD FirstName \* Upper

</td>
    <td>«FIRSTNAME»

</td>
  </tr>
</table>
TEST SEQ

<table>
  <tr>
    <td>SEQ Figure

</td>
    <td>1

</td>
  </tr>
  <tr>
    <td>SEQ Figure

</td>
    <td>2

</td>
  </tr>
  <tr>
    <td>SEQ Figure \c

</td>
    <td>2

</td>
  </tr>
  <tr>
    <td>SEQ Figure \r 5

</td>
    <td>5

</td>
  </tr>
  <tr>
    <td>SEQ Figure \h

</td>
    <td>

</td>
  </tr>
  <tr>
    <td>SEQ Figure \h \* Arabic

</td>
    <td>7

</td>
  </tr>
  <tr>
    <td>SEQ Figure \s 1

</td>
    <td>8

</td>
  </tr>
</table>