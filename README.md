# flatxml2excel

A simple and fast library to create MS Office Excel(>2007) build on top of the followings modules;

- [msexcel-builder](http://www.github.com/chuanyi/msexcel-builder.git). 
- [xml2json](http://www.github.com/buglabs/node-xml2json)

The codes to flatten nested xml are adopted from [flatjson](github.com/freezer333/flatjson) 

## Installation
```
npm install flatxml2excel
```

## Usage
- Default usage result to creation of excel file called sample.xlsx 
```
flatxml2excel = require('flatxml2excel');
sample = '<hello>Mambo</hello>'
 //will produce excel file called sample.xlsx in the current directoty
flatxml2excel.buildXls(sample, function(ok){
    if(ok!==null){
        throw 'fails';
    }
})
```

- Using options you can specify the output excel filepath and filename
```
flatxml2excel = require('flatxml2excel');
sample = '<hello>Mambo</hello>'
options = {
    filepath:'./',
    filename:'output.xlsx'
}
 //will produce excel file called output.xlsx in the filepath specified
flatxml2excel.buildXls(sample, options, function(ok){
    if(ok!==null){
        throw 'fails';
    }
})
```

- Using options you can also produce buffer output instead of excel file
```
flatxml2excel = require('flatxml2excel');
sample = '<hello>Mambo</hello>'
options = {
  buffer:true
}
 //will produce buffer output
flatxml2excel.buildXls(sample, options, function(data){
    if(data===null){
        throw 'fails';
    }
    return data;
})
```

- You can also use this library as middleware to download excel direct using browser
```
var fs = require('fs');
var express = require('express');
var app = express();
flatxml2excel = require('flatxml2excel');
app.use(flatxml2excel.middleware);

app.get('/download/sample.xlsx',function(req, res){
    res.xls(data.xlsx, fs.readFileSync(/path/to/excelfile, 'utf8'));
    })
```


## Core Features;

* Produce **multiple worksheet** for nested XML file.

```
var test = {
    parent1: {
        child1: {
            child11: {
                child111: 10,
                child112: 14,
                child1111: {
                    child1112: 15
                }
            },
            child12: 11
        },
        child2: {
            child21: 20,
            child22: 21
        }
    },
    parent2: {
        child3: {
            child11: 101,
            child21: 102
        }
    }
}
```
**produce**

![sample image]
(file:///tmp/tmp0uf5d0.html/sample.png)

* **coerce:** Makes type coercion. i.e.: numbers and booleans present in attributes and element values are converted from string to its correspondent data types. Coerce can be optionally defined as an object with specific methods of coercion based on attribute name or tag name, with fallback to default coercion.
* **sanitize:** Sanitizes the following characters present in element values:

```javascript
var chars =  {
    '<': '&lt;',
    '>': '&gt;',
    '(': '&#40;',
    ')': '&#41;',
    '#': '&#35;',
    '&': '&amp;',
    '"': '&quot;',
    "'": '&apos;'
};
```


## License
(The MIT License)

Copyright 2015 byteskode. All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to
deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.