

/*
*   Export excel file from Html table.
*
*   - Required
*       jquery: jquery-1.12.x.js or later
*       https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.9.10/xlsx.full.min.js
*       https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.3/FileSaver.min.js
*
*   - Usage
*    $.toExcel("output.xlsx", [{sheetName: "sheet1", selector: "#tableselector"},...], option);
*/
(function($) {
    $.extend(
        {
            tableToExcel: function(filename, data, option={}) {
                if (data && data.length > 0) {
                    var defaults = {
                        bookType: 'xlsx',
                        bookSST: false,
                        type: 'binary'
                    };
                    var option = $.extend(defaults, option);
                    var wb = {
                        SheetNames: [],
                        Sheets: {}
                    };

                    $.each(data, function(k,v) {
                        wb.SheetNames[k] = v.sheetName;
                        wb.Sheets[v.sheetName] = XLSX.utils.table_to_sheet($(v.selector)[0], option);
                    });

                    var wbout = XLSX.write(wb, option);
                    var toBuffer = function(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) {
                            view[i] = s.charCodeAt(i) & 0xFF;
                        }
                        return buf;
                    }
                    saveAs(new Blob([toBuffer(wbout)], {type: 'application/octet-stream'}), filename);
                }
            }
        }
    )
})(jQuery);
