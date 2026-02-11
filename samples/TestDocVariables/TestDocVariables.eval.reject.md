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
    <td style="border:0.5pt solid #000000;">two

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCVARIABLE Missing

</td>
    <td style="border:0.5pt solid #000000;">Error! No document variable supplied.

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;"><b>DOCVARIABLE</b> Var1 \* Charformat

</td>
    <td style="border:0.5pt solid #000000;"><b>two</b>

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCVARIABLE Var1 \* MERGEFORMAT

</td>
    <td style="border:0.5pt solid #000000;"><u>two</u>

</td>
  </tr>
</table>
<p style="text-align:center;">IF TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">IF 5 >= 3 "OK" "BAD"

</td>
    <td style="border:0.5pt solid #000000;">OK

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">IF "Approved" = "App*" "YES" "NO"

</td>
    <td style="border:0.5pt solid #000000;">YES

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SET Var2 "two"

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">IF { REF Var1 } = "two" "Value is { REF Var1 }" "Mismatch"

</td>
    <td style="border:0.5pt solid #000000;">Value is two

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">IF { REF MissingBookmark } = "" "Empty" "Not Empty"

</td>
    <td style="border:0.5pt solid #000000;">Not Empty

</td>
  </tr>
</table>
<p style="text-align:center;">DOCPROPERTY TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">DOCPROPERTY Title

</td>
    <td style="border:0.5pt solid #000000;">Field Test Title

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCPROPERTY Created \@ "yyyy-MM-dd"

</td>
    <td style="border:0.5pt solid #000000;">2010-10-13

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">DOCPROPERTY "CustomProp1"

</td>
    <td style="border:0.5pt solid #000000;">custom-value

</td>
  </tr>
</table>
<p style="text-align:center;">TEST COMPARE</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">COMPARE 5 >= 3

</td>
    <td style="border:0.5pt solid #000000;">1

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">COMPARE 2 > 3

</td>
    <td style="border:0.5pt solid #000000;">0

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">COMPARE "Approved" = "App*"

</td>
    <td style="border:0.5pt solid #000000;">1

</td>
  </tr>
</table>
<p style="text-align:center;">TEST MERGEFIELD</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD FirstName \b "Hello " \f "!"

</td>
    <td style="border:0.5pt solid #000000;">Hello Ana!

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD EmptyField \b "Hello " \f "!"

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD GivenName \m

</td>
    <td style="border:0.5pt solid #000000;">Ana

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD FirstName \v

</td>
    <td style="border:0.5pt solid #000000;">A
n
a

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">MERGEFIELD FirstName \* Upper

</td>
    <td style="border:0.5pt solid #000000;">ANA

</td>
  </tr>
</table>
<p style="text-align:center;">TEST SEQ</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure

</td>
    <td style="border:0.5pt solid #000000;">1

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure

</td>
    <td style="border:0.5pt solid #000000;">2

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \c

</td>
    <td style="border:0.5pt solid #000000;">2

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \r 5

</td>
    <td style="border:0.5pt solid #000000;">5

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \h

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \h \* Arabic

</td>
    <td style="border:0.5pt solid #000000;">7

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SEQ Figure \s 1

</td>
    <td style="border:0.5pt solid #000000;">8

</td>
  </tr>
</table>
<p style="text-align:center;">TEST FORMULA</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">= 2 + 3

</td>
    <td style="border:0.5pt solid #000000;">5

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= 2 + 3 * 4

</td>
    <td style="border:0.5pt solid #000000;">14

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= 3 > 2

</td>
    <td style="border:0.5pt solid #000000;">1

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= SUM(1,2,3)

</td>
    <td style="border:0.5pt solid #000000;">6

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= { DATE \@ "yyyy" } + 1

</td>
    <td style="border:0.5pt solid #000000;">2027

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= 1 / 0

</td>
    <td style="border:0.5pt solid #000000;">!Zero Divide

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= LN(10)

</td>
    <td style="border:0.5pt solid #000000;">!Syntax Error, 10

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">= SUM(ABOVE)

</td>
    <td style="border:0.5pt solid #000000;">2053

</td>
  </tr>
</table>
<p style="text-align:center;">ASK TEST</p>


<table style="border:0.5pt solid #000000;border-collapse:collapse;">
  <tr>
    <td style="border:0.5pt solid #000000;">ASK Name "Name?" \d "Unknown"

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">REF Name

</td>
    <td style="border:0.5pt solid #000000;">Bob

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">ASK Name "Name?" \o

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">REF Name

</td>
    <td style="border:0.5pt solid #000000;">Unknown

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SET DefaultCity "Rome"

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">SET Greeting "Hi"

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">ASK City "{ REF Greeting }?" \d "{ REF DefaultCity }"

</td>
    <td style="border:0.5pt solid #000000;">

</td>
  </tr>
  <tr>
    <td style="border:0.5pt solid #000000;">REF City

</td>
    <td style="border:0.5pt solid #000000;">Montreal

</td>
  </tr>
</table>
</div>
</div>