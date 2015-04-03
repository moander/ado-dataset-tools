This is a library of tools that has grown from a few functions for manipulating ADODB **recordsets** for VB and VBA to an implementation in C# for ADO.net **datatable** and **dataset**.

The library includes functions for;
  * Group By
  * Sub totals
  * Pivots (Cross Tabs)
  * Cloning
  * Merging
  * Appending


of datatables and recordsets.

It's original use was for manipulating datasets that didn't originate from a database so all aggregate functions had to be done in code.

Initially i'm tidying, documenting and uploading the ADODB version, the ADO.net version will follow soon. Mostly this is so i can get the hang of Google Code :) Before i get to make any packages to download you can browse and download the source from [SVN](http://code.google.com/p/ado-dataset-tools/source/browse/#svn/trunk/VB)

**Update:**
You can download all VB code plus a simple set of test from the [examples in SVN](http://ado-dataset-tools.googlecode.com/svn/trunk/VB/examples/ado-recordset-unit-tests.xls).

Once copied, try `RunAllTests()` in module `LocalUnitTests`

Ive also included function `GetGoogleFinanceData()`

eg:

`Dim rs as ADODB.Recordset`

`set rs = GetGoogleFinanceData("NASDAQ:GOOG",startDate,endDate)`



This is just to get by until a proper release.

m