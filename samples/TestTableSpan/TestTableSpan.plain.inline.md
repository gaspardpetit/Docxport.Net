<!-- Author: Petit, Gaspard -->
<!-- LastModifiedBy: Petit, Gaspard -->
<!-- Revision: 4 -->
<!-- Created: 2026-01-09 02:37:00Z -->
<!-- Modified: 2026-01-09 02:50:00Z -->

Normal

<table>
  <tr>
    <td>1,1

</td>
    <td>1,2

</td>
    <td>1,3

</td>
  </tr>
  <tr>
    <td>2,1-2,2

</td>
    <td></td>
    <td>2,3

</td>
  </tr>
  <tr>
    <td>3,1

</td>
    <td>3,2

</td>
    <td>3,3

</td>
  </tr>
</table>
Col Span

<table>
  <tr>
    <td>1,1

</td>
    <td>1,2

</td>
    <td>1,3

</td>
  </tr>
  <tr>
    <td colspan="2">2,1-2,2

</td>
    <td>2,3

</td>
  </tr>
  <tr>
    <td>3,1

</td>
    <td>3,2

</td>
    <td>3,3

</td>
  </tr>
</table>
Row Span

<table>
  <tr>
    <td>1,1

</td>
    <td>1,2

</td>
    <td>1,3

</td>
  </tr>
  <tr>
    <td>2,1-

</td>
    <td rowspan="2">2,2 + 3,2

</td>
    <td>2,3

</td>
  </tr>
  <tr>
    <td>3,1

</td>
    <td>3,3

</td>
  </tr>
</table>
Row Span

<table>
  <tr>
    <td>1,1

</td>
    <td>1,2

</td>
    <td>1,3

</td>
  </tr>
  <tr>
    <td>2,1-

</td>
    <td rowspan="2" colspan="2">2,2 + 3,2 + 2,3 + 3,3

</td>
  </tr>
  <tr>
    <td>3,1

</td>
  </tr>
</table>