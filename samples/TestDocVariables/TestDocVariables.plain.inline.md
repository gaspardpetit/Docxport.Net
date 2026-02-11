<!-- Title: Field Test Title -->
<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 14 -->
<!-- Created: 2026-02-11 08:11:00Z -->
<!-- Modified: 2026-02-11 09:59:00Z -->
<!-- CustomProp1: custom-value -->
<!-- Created: 2010-10-13T04:00:00Z -->

DOCVARIABLE TEST

<table>
  <tr>
    <td>DOCVARIABLE Var1

</td>
    <td>

`DOCVARIABLE Var1`

one

</td>
  </tr>
  <tr>
    <td>DOCVARIABLE Missing

</td>
    <td>

`DOCVARIABLE Missing`

<b>Error! No document variable supplied.</b>

</td>
  </tr>
  <tr>
    <td><b>DOCVARIABLE</b> Var1 \* Charformat

</td>
    <td>

`DOCVARIABLE Var1 \* Charformat`

<b>one</b>

</td>
  </tr>
  <tr>
    <td>DOCVARIABLE Var1 \* MERGEFORMAT

</td>
    <td>

`DOCVARIABLE Var1 \* MERGEFORMAT`

<u>one</u>

</td>
  </tr>
</table>
IF TEST

<table>
  <tr>
    <td>IF 5 >= 3 "OK" "BAD"

</td>
    <td>

`IF 5 >= 3 "OK" "BAD"`

OK

</td>
  </tr>
  <tr>
    <td>IF "Approved" = "App*" "YES" "NO"

</td>
    <td>

`IF "Approved" = "App*" "YES" "NO"`

YES

</td>
  </tr>
  <tr>
    <td>SET Var2 "two"

</td>
    <td>

`SET Var2 "two"`

two

</td>
  </tr>
  <tr>
    <td>IF { REF Var1 } = "two" "Value is { REF Var1 }" "Mismatch"

</td>
    <td>

`IF`



`REF Var`



`2`



`two`



`= "two" "Value is`



`REF Var`



`2`



`two`



`" "Mismatch"`

Value is two

</td>
  </tr>
  <tr>
    <td>IF { REF MissingBookmark } = "" "Empty" "Not Empty"

</td>
    <td>

`IF { REF MissingBookmark } = "" "Empty" "Not Empty"`

REF

</td>
  </tr>
</table>
DOCPROPERTY TEST

<table>
  <tr>
    <td>DOCPROPERTY Title

</td>
    <td>

`DOCPROPERTY Title`

Field Test Title

</td>
  </tr>
  <tr>
    <td>DOCPROPERTY Created \@ "yyyy-MM-dd"

</td>
    <td>

`DOCPROPERTY Created \@ "yyyy-MM-dd"`

2010-10-13

</td>
  </tr>
  <tr>
    <td>DOCPROPERTY "CustomProp1"

</td>
    <td>

`DOCPROPERTY "CustomProp1"`

custom-value

</td>
  </tr>
</table>
TEST COMPARE

<table>
  <tr>
    <td>COMPARE 5 >= 3

</td>
    <td>

`COMPARE 5 >= 3`

1

</td>
  </tr>
  <tr>
    <td>COMPARE 2 > 3

</td>
    <td>

`COMPARE 2 > 3`

0

</td>
  </tr>
  <tr>
    <td>COMPARE "Approved" = "App*"

</td>
    <td>

`COMPARE "Approved" = "App*"`

1

</td>
  </tr>
</table>
TEST MERGEFIELD

<table>
  <tr>
    <td>MERGEFIELD FirstName \b "Hello " \f "!"

</td>
    <td>

`MERGEFIELD FirstName \b "Hello " \f "!"`

Hello «FirstName»!

</td>
  </tr>
  <tr>
    <td>MERGEFIELD EmptyField \b "Hello " \f "!"

</td>
    <td>

`MERGEFIELD EmptyField \b "Hello " \f "!"`

Hello «EmptyField»!

</td>
  </tr>
  <tr>
    <td>MERGEFIELD GivenName \m

</td>
    <td>

`MERGEFIELD GivenName \m`

«GivenName»

</td>
  </tr>
  <tr>
    <td>MERGEFIELD FirstName \v

</td>
    <td>

`MERGEFIELD FirstName \v`

«FirstName»

</td>
  </tr>
  <tr>
    <td>MERGEFIELD FirstName \* Upper

</td>
    <td>

`MERGEFIELD FirstName \* Upper`

«FIRSTNAME»

</td>
  </tr>
</table>
TEST SEQ

<table>
  <tr>
    <td>SEQ Figure

</td>
    <td>

`SEQ Figure`

1

</td>
  </tr>
  <tr>
    <td>SEQ Figure

</td>
    <td>

`SEQ Figure`

2

</td>
  </tr>
  <tr>
    <td>SEQ Figure \c

</td>
    <td>

`SEQ Figure \c`

2

</td>
  </tr>
  <tr>
    <td>SEQ Figure \r 5

</td>
    <td>

`SEQ Figure \r 5`

5

</td>
  </tr>
  <tr>
    <td>SEQ Figure \h

</td>
    <td>

`SEQ Figure \h`



</td>
  </tr>
  <tr>
    <td>SEQ Figure \h \* Arabic

</td>
    <td>

`SEQ Figure \h \* Arabic`

7

</td>
  </tr>
  <tr>
    <td>SEQ Figure \s 1

</td>
    <td>

`SEQ Figure \s 1`

8

</td>
  </tr>
</table>
TEST FORMULA

<table>
  <tr>
    <td>= 2 + 3

</td>
    <td>

`= 2 + 3`

5

</td>
  </tr>
  <tr>
    <td>= 2 + 3 * 4

</td>
    <td>

`= 2 + 3 * 4`

14

</td>
  </tr>
  <tr>
    <td>= 3 > 2

</td>
    <td>

`= 3 > 2`

1

</td>
  </tr>
  <tr>
    <td>= SUM(1,2,3)

</td>
    <td>

`= SUM(1,2,3)`

6

</td>
  </tr>
  <tr>
    <td>= { DATE \@ "yyyy" } + 1

</td>
    <td>

`=`



`DATE \@ "yyyy"`



`2026`



`+ 1`

2027

</td>
  </tr>
  <tr>
    <td>= 1 / 0

</td>
    <td>

`= 1 / 0`

<b>!Zero Divide</b>

</td>
  </tr>
  <tr>
    <td>= LN(10)

</td>
    <td>

`= LN(10)`

<b>!Syntax Error, 10</b>

</td>
  </tr>
  <tr>
    <td>= SUM(ABOVE)

</td>
    <td>

`= SUM(ABOVE)`

2053

</td>
  </tr>
</table>
ASK TEST

<table>
  <tr>
    <td>ASK Name "Name?" \d "Unknown"

</td>
    <td>

`ASK Name "Name?" \d "Unknown"`

Bob

</td>
  </tr>
  <tr>
    <td>REF Name

</td>
    <td>

`REF Name`

Bob

</td>
  </tr>
  <tr>
    <td>ASK Name "Name?" \o

</td>
    <td>

`ASK Name "Name?" \o`

Bob

</td>
  </tr>
  <tr>
    <td>REF Name

</td>
    <td>

`REF Name`

Bob

</td>
  </tr>
  <tr>
    <td>SET DefaultCity "Rome"

</td>
    <td>

`SET DefaultCity "Rome"`

Rome

</td>
  </tr>
  <tr>
    <td>SET Greeting "Hi"

</td>
    <td>

`SET Greeting "Hi"`

Hi

</td>
  </tr>
  <tr>
    <td>ASK City "{ REF Greeting }?" \d "{ REF DefaultCity }"

</td>
    <td>

`ASK City "`



`REF Greeting`



`Hi`



`REF`



`Name`



`Bob`



`?" \d "`



`REF DefaultCity`



`Rome`



`"`

Montreal

</td>
  </tr>
  <tr>
    <td>REF City

</td>
    <td>

`REF`



`City`

Montreal

</td>
  </tr>
</table>