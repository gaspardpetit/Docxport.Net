<!-- Title: Field Test Title -->
<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 14 -->
<!-- Created: 2026-02-11 08:11:00Z -->
<!-- Modified: 2026-02-11 09:59:00Z -->
<!-- CustomProp1: custom-value -->
<!-- Created: 2010-10-13T04:00:00Z -->

<div class="section" style="color:#000000;display:flex;flex-direction:column;position:relative;width:8.5in;min-height:11in;box-sizing:border-box;padding-left:1in;padding-right:1in;background-color:#ffffff;font-family:Aptos;font-size:12pt;">
<div class="body" style="flex:1 0 auto;padding-top:1in;">
<p style="text-align:center;">DOCVARIABLE TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">DOCVARIABLE Var1

</td>
    <td style="border:0.5pt solid #000000;">

`DOCVARIABLE Var1`

one

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCVARIABLE Missing

</td>
    <td style="border:0.5pt solid #000000;">

`DOCVARIABLE Missing`

<b>Error! No document variable supplied.</b>

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;"><b>DOCVARIABLE</b> Var1 \* Charformat

</td>
    <td style="border:0.5pt solid #000000;">

`DOCVARIABLE Var1 \* Charformat`

<b>one</b>

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCVARIABLE Var1 \* MERGEFORMAT

</td>
    <td style="border:0.5pt solid #000000;">

`DOCVARIABLE Var1 \* MERGEFORMAT`

<u>one</u>

</td>
  </tr>
</table>
<p style="text-align:center;">IF TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">IF 5 >= 3 "OK" "BAD"

</td>
    <td style="border:0.5pt solid #000000;">

`IF 5 >= 3 "OK" "BAD"`

OK

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">IF "Approved" = "App*" "YES" "NO"

</td>
    <td style="border:0.5pt solid #000000;">

`IF "Approved" = "App*" "YES" "NO"`

YES

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SET Var2 "two"

</td>
    <td style="border:0.5pt solid #000000;">

`SET Var2 "two"`

<a id="Var2" data-bookmark-id="0"></a>two

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">IF { REF Var1 } = "two" "Value is { REF Var1 }" "Mismatch"

</td>
    <td style="border:0.5pt solid #000000;">

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
    <td style="border:0.5pt solid #000000;">IF { REF MissingBookmark } = "" "Empty" "Not Empty"

</td>
    <td style="border:0.5pt solid #000000;">

`IF { REF MissingBookmark } = "" "Empty" "Not Empty"`

REF

</td>
  </tr>
</table>
<p style="text-align:center;">DOCPROPERTY TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">DOCPROPERTY Title

</td>
    <td style="border:0.5pt solid #000000;">

`DOCPROPERTY Title`

Field Test Title

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCPROPERTY Created \@ "yyyy-MM-dd"

</td>
    <td style="border:0.5pt solid #000000;">

`DOCPROPERTY Created \@ "yyyy-MM-dd"`

2010-10-13

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCPROPERTY "CustomProp1"

</td>
    <td style="border:0.5pt solid #000000;">

`DOCPROPERTY "CustomProp1"`

custom-value

</td>
  </tr>
</table>
<p style="text-align:center;">TEST COMPARE</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">COMPARE 5 >= 3

</td>
    <td style="border:0.5pt solid #000000;">

`COMPARE 5 >= 3`

1

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">COMPARE 2 > 3

</td>
    <td style="border:0.5pt solid #000000;">

`COMPARE 2 > 3`

0

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">COMPARE "Approved" = "App*"

</td>
    <td style="border:0.5pt solid #000000;">

`COMPARE "Approved" = "App*"`

1

</td>
  </tr>
</table>
<p style="text-align:center;">TEST MERGEFIELD</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD FirstName \b "Hello " \f "!"

</td>
    <td style="border:0.5pt solid #000000;">

`MERGEFIELD FirstName \b "Hello " \f "!"`

Hello «FirstName»!

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD EmptyField \b "Hello " \f "!"

</td>
    <td style="border:0.5pt solid #000000;">

`MERGEFIELD EmptyField \b "Hello " \f "!"`

Hello «EmptyField»!

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD GivenName \m

</td>
    <td style="border:0.5pt solid #000000;">

`MERGEFIELD GivenName \m`

«GivenName»

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD FirstName \v

</td>
    <td style="border:0.5pt solid #000000;">

`MERGEFIELD FirstName \v`

«FirstName»

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD FirstName \* Upper

</td>
    <td style="border:0.5pt solid #000000;">

`MERGEFIELD FirstName \* Upper`

«FIRSTNAME»

</td>
  </tr>
</table>
<p style="text-align:center;">TEST SEQ</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure`

1

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure`

2

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \c

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure \c`

2

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \r 5

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure \r 5`

5

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \h

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure \h`



</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \h \* Arabic

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure \h \* Arabic`

7

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \s 1

</td>
    <td style="border:0.5pt solid #000000;">

`SEQ Figure \s 1`

8

</td>
  </tr>
</table>
<p style="text-align:center;">TEST FORMULA</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">= 2 + 3

</td>
    <td style="border:0.5pt solid #000000;">

`= 2 + 3`

5

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= 2 + 3 * 4

</td>
    <td style="border:0.5pt solid #000000;">

`= 2 + 3 * 4`

14

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= 3 > 2

</td>
    <td style="border:0.5pt solid #000000;">

`= 3 > 2`

1

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= SUM(1,2,3)

</td>
    <td style="border:0.5pt solid #000000;">

`= SUM(1,2,3)`

6

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= { DATE \@ "yyyy" } + 1

</td>
    <td style="border:0.5pt solid #000000;">

`=`



`DATE \@ "yyyy"`



`2026`



`+ 1`

2027

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= 1 / 0

</td>
    <td style="border:0.5pt solid #000000;">

`= 1 / 0`

<b>!Zero Divide</b>

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= LN(10)

</td>
    <td style="border:0.5pt solid #000000;">

`= LN(10)`

<b>!Syntax Error, 10</b>

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= SUM(ABOVE)

</td>
    <td style="border:0.5pt solid #000000;">

`= SUM(ABOVE)`

2053

</td>
  </tr>
</table>
<p style="text-align:center;">ASK TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">ASK Name "Name?" \d "Unknown"

</td>
    <td style="border:0.5pt solid #000000;">

`ASK Name "Name?" \d "Unknown"`

Bob

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">REF Name

</td>
    <td style="border:0.5pt solid #000000;">

`REF Name`

Bob

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">ASK Name "Name?" \o

</td>
    <td style="border:0.5pt solid #000000;">

`ASK Name "Name?" \o`

<a id="Name" data-bookmark-id="1"></a>Bob

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">REF Name

</td>
    <td style="border:0.5pt solid #000000;">

`REF Name`

Bob

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SET DefaultCity "Rome"

</td>
    <td style="border:0.5pt solid #000000;">

`SET DefaultCity "Rome"`

<a id="DefaultCity" data-bookmark-id="2"></a>Rome

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SET Greeting "Hi"

</td>
    <td style="border:0.5pt solid #000000;">

`SET Greeting "Hi"`

<a id="Greeting" data-bookmark-id="3"></a>Hi

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">ASK City "{ REF Greeting }?" \d "{ REF DefaultCity }"

</td>
    <td style="border:0.5pt solid #000000;">

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

<a id="City" data-bookmark-id="4"></a>Montreal

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">REF City

</td>
    <td style="border:0.5pt solid #000000;">

`REF`



`City`

Montreal