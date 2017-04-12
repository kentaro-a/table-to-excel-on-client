# table-to-excel-on-client

## Overview
Export a excel file with multiple sheets from Html table on just only client side.  
The datasource of excel is ordinary table tag structured by thead,tbody,tr,td.  

##Required
- jquery-1.12.x.js or later
- https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.9.10/xlsx.full.min.js
- https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.3/FileSaver.min.js


##Usage
This plugin has Extended JqueryObject($) to bind the plugin method, so type the codes like this when you use.  
``
$.toExcel("output.xlsx", [{sheetName: "sheet1", selector: "#tableselector"},...], option);
``
