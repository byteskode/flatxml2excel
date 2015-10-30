var excelbuilder = require('msexcel-builder');
var parser = require('xml2json');
var fs = require('fs');


var flattenJSON = function(json, filter) {
    var flattened = {};

    if (!filter) {
        filter = function(key, item) {
            return !(typeof item === 'function');
        }
    }

    walk(json);

    return flattened;
    /**
     * [walk description]
     * Walk through JSON object to prepare object to add to root JSON obj2convert flattened
     * @param  {Object} obj2convert  JSON obj2convert to walk through
     * @param  {String} path key associated with JSON obj2convert to walk through
     */
    function walk(obj2convert, path) {
        var item;
        var obj = {};
        var objectValues = [];
        path = path || "";
        for (var key in obj2convert) {
            item = obj2convert[key];
            if (obj2convert.hasOwnProperty(key) && filter(key, item)) {
                if (typeof item === 'object') {
                    var map = new Object();
                    map[key] = item;
                    objectValues.push(map);
                } else {
                    obj[key] = item;
                }
            }
        }
        if (Object.keys(obj).length !== 0) {
            if (Object.keys(flattened).length === 0 && objectValues.length === 0) {
                flattened = obj;
            } else {
                flattened[path] = obj;
            };

        }
        objectValues.forEach(function(obj) {
            var key = Object.keys(obj)[0];
            walk(obj[key], key);
        })
    }
}

var buildXls = function(xml, options, callback) {
    if (typeof arguments[arguments.length - 1] !== 'function') {
        throw 'no callback function';
    }

    if (typeof options === 'function') {
        callback = options;
        options = {};
        options.buffer = false;
    };

    var workbook;
    if (options.hasOwnProperty('filepath') && options.hasOwnProperty('filename')) {
        workbook = excelbuilder.createWorkbook(options.filepath, options.filename);
    } else {
        var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx');
    }
    
    var result = parser.toJson(xml, {
        object: true,
        coerce: true,
        trim: true,
        sanitize: true
    });
    var flatJson = flattenJSON(result);
    var multipleSheet = false;
    for (var key in flatJson) {
        var item = flatJson[key];
        if (typeof item === 'object') {
            multipleSheet = true;
            buildSheet(workbook, item, key);
        }
    }
    if (!multipleSheet) {
        buildSheet(workbook, flatJson);
    }

    // Save it
    workbook.save(function(ok) {
        if (!options.buffer) {
            callback(ok);
        } else {
            if (ok !== null) {
                //error so pass null
                callback(null);
            };
            fs.readFile('./sample.xlsx', function(err, data) {
                if (err) {
                    throw err;
                };
                //pass buffer object
                callback(data);
                fs.unlinkSync('./sample.xlsx');
            })
        }
    });


    function buildSheet(workbook, sheetData, sheetName) {
        var allKeys = Object.keys(sheetData);
        var keyCount = allKeys.length;
        var sheetName = sheetName ? sheetName : 'sheet';
        //increment row number to avoid scenario when the sheet has single key(column)
        var sheet = workbook.createSheet(sheetName, keyCount, keyCount + 1);
        for (var i = 0; i < keyCount; i++) {
            var key = allKeys[i];
            sheet.set(i + 1, 1, key);
            sheet.set(i + 1, 2, sheetData[key]);
        };
    }
}


module.exports = {

    flatten: flattenJSON,

    buildXls: buildXls,

    middleware: function(req, res, next) {
        res.xls = function(fn, xml) {
            var xls = buildXls(xml, function(data) {
                res.setHeader('Content-Type', 'application/vnd.openxmlformats');
                res.setHeader("Content-Disposition", "attachment; filename=" + fn);
                res.end(data, 'binary');
            })
        };
        next();
    }
}
