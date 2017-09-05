# PyMatcherThis python script was written to merge data of 2 Excel files with a match on 2 columns.# Installation<pre>pip install openpyxlpip install xlrdpip install xlwtgit clone https://github.com/Sylvaner/PyMatcher</pre># Usage## Sample input data### input1.xlsx<table>  <tr>    <th>Lastname</th>    <th>Firstname</th>    <th>ID</th>  </tr>  <tr>    <td>Peregrin</td>    <td>Touc</td>    <td>3434342</td>  </tr>  <tr>    <td>brandibouc</td>    <td>meriadoc</td>    <td>2369127</td>  </tr>  <tr>    <td>Chaumine</td>    <td>Rose</td>    <td>320988</td>  </tr>  <tr>    <td>Sacquet</td>    <td>bilbo</td>    <td>239820</td>  </tr>  <tr>    <td>SACQUET</td>    <td>FRODO</td>    <td>29399</td>  </tr>  <tr>    <td>BOLGEURRE</td>    <td>Estella</td>    <td>238927</td>  </tr></table>### input2.xls<table>  <tr>    <th>ID</th>    <th>Gender</th>  </tr>  <tr>    <td>320988</td>    <td>Female</td>  </tr>  <tr>    <td>239820</td>    <td>Male</td>  </tr>  <tr>    <td>3434342</td>    <td>Male</td>  </tr>  <tr>    <td>29399</td>    <td>Male</td>  </tr>  <tr>    <td>2369127</td>    <td>Male</td>  </tr></table>## Merge on ID```python pymatcher.py input1.xls 3 input2.xls 1 output.xls```### Result<table>  <tr>    <th>ID</th>    <th>Lastname</th>    <th>Firstname</th>    <th>Gender</th>  </tr>  <tr>    <td>3434342</td>    <td>Peregrin</td>    <td>Touc</td>    <td>Male</td>  </tr>  <tr>    <td>2369127</td>    <td>brandibouc</td>    <td>meriadoc</td>    <td>Male</td>  </tr>  <tr>    <td>320988</td>    <td>Chaumine</td>    <td>Rose</td>    <td>Female</td>  </tr>  <tr>    <td>239820</td>    <td>Sacquet</td>    <td>bilbo</td>    <td>Male</td>  </tr>  <tr>    <td>29399</td>    <td>SACQUET</td>    <td>FRODO</td>    <td>Male</td>  </tr></table># Options## --no-headerDon't process the first line as header.## --output-sheetname=NAMESet the name of the sheet in the output spreadsheet.