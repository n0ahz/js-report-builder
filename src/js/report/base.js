function BaseReport() {
    /* The root class of any report */
    this.headers = [];
    this.rows = [];

    this.addRow = function(d, isHeader) {
        /* add a row, specifying if its a header or a data row */
        if (isHeader) {
            this.headers.push(d);
        } else {
            this.rows.push(d);
        }
    };

    this.getData = function() {
        /* get the full data of the report */
        return _.concat(this.headers, this.rows);
    };
}

function ExcelReport() {
    /*
    extends the BaseReport class; this has some specific properties and methods for ms-excel report
    the formatting and the methods depend on the 3rd party libs:
     https://github.com/gitbrent/xlsx-js-style which extends https://sheetjs.com/ and provides support for styles.
     Also depends on lodash for utility works.
    */
    BaseReport.call(this);
    this.mergeSettings = [];
    this.colWidths = [];

    this.prepareStyle = function(settings) {
        /* prepare the styles object from an external provided flattened setting.
        * the final format is a nested setting object required by the extended XLSX plugin.
        * not all formatting are taken now..can be extended later as required.
        */
        var styles = {alignment: {vertical: settings.vAlign, horizontal: settings.hAlign}};
        if (settings.fillFgColor) {
            styles['fill'] = {fgColor: {rgb: settings.fillFgColor}};
        }
        if (settings.bold) {
            styles['font'] = {bold: true};
        }
        return styles;
    };

    this.wrapWithStyle = function (value, options) {
        /* wrap a cell data value with styles and format it as required by XLSX plugin */
        var settings = {
            type: 't',
            vAlign: 'top',
            hAlign: 'left',
            fillFgColor: '',
            bold: false
        };
        _.merge(settings, options);
        return {v: value, t: settings.type, s: this.prepareStyle(settings)};
    };

    this.modifyStyle = function(data, styleSetting) {
        /* modify a particular cell styles with the provided new styleSetting */
        var oldStyle = data['s'];
        _.merge(oldStyle, this.prepareStyle(styleSetting));
    };

    this.setAllColWidths = function(width) {
        /* can be used to specify widths of all cells by a singular value */
        var self = this;
        _.times(this.rows[0].length, function(n) {
            self.colWidths.push({width: width});
        });
    };

    this.applyMerge = function(startRow, startCol, endRow, endCol) {
        /* prepare and add a merge setting with the range specified by the arguments */
        this.mergeSettings.push({s: {r: startRow, c: startCol}, e: {r: endRow, c: endCol}});
    };

    this.mergeSameValueOfRows = function(columnIndices) {
        /* can be used to merge the rows of a single column vertically for the range where the values are same */
        var self = this;
        columnIndices.forEach(function(ci, i) {
            var rowSpan = 0;
            var previousValue = '';
            var startIndex = 0;
            var rowOffset = self.headers.length;
            self.rows.forEach(function(d, index) {
                var currentValue = d[ci]['v'];
                if (previousValue) {
                    if (currentValue === previousValue) {
                        rowSpan += 1;
                    } else {
                        if (rowSpan) {
                            self.applyMerge(startIndex, ci, startIndex + rowSpan, ci);
                            rowSpan = 0;
                        }
                        startIndex = index + rowOffset;
                    }
                } else {
                    startIndex = index + rowOffset;
                }
                previousValue = currentValue;
            });

            if (rowSpan) {
                self.applyMerge(startIndex, ci, startIndex + rowSpan, ci);
                rowSpan = 0;
            }
        });
    };

    this.fitToColumn = function() {
        /* get maximum character of each column based on all data and values and prepare the best size of the columns.
        * kept a minimum width and a gutter value for columns with empty data.
        */
        var colWidths = [];
        var minColWidth = 7;
        this.rows.forEach(function(r, i) {
            r.forEach(function(d, di) {
                var dataLength = d['v'].toString().length;
                colWidths[di] = colWidths[di] !== undefined ? Math.max(colWidths[di], dataLength, minColWidth) : dataLength;
            });
        });
        return colWidths.map(function(cw) { return {width: cw + 2}; });
    };

    this.generateFile = function(fileName) {
        /* generate and export the xlsx file with the provided filename.
        * when no fileName is provided, it will have a default name.
        * did not provide the argument for modifying the sheet name. currently there will be only 1 sheet named: report
        */
        var ext = '.xlsx';
        if (!fileName) fileName = 'generated-excel-file';
        if (!fileName.endsWith(ext)) fileName += ext;
        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet(this.getData());
        ws["!merges"] = this.mergeSettings;
        ws["!cols"] = this.colWidths.length > 0 ? this.colWidths : this.fitToColumn(this.rows);
        XLSX.utils.book_append_sheet(wb, ws, 'report');
        XLSX.writeFile(wb, fileName);
    };
}

ExcelReport.prototype = Object.create(Base.prototype);
