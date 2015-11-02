expect = require('chai').expect;
should = require('chai').should;
var parser = require('xml2json');
var json2xls = require('json2xls');
xmlXlsTool = require('../xmlXlsTool');

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

var xmlData = '<?xml version="1.0" encoding="utf-8"?><root><parent1><child1><child11><child111>10</child111>' +
    '<child112>14</child112><child1111><child1112>15</child1112></child1111>' +
    '</child11><child12>12</child12></child1><child2><child21>20</child21>' +
    '<child22>21</child22></child2></parent1>' +
    '<parent2><child3><child11>101</child11><child21>102</child21></child3></parent2>' +
    '</root>';

var flatJson = {
    parent1: 11,
    parent2: 12,
    parent3: 13,
}

var xmlFlatData = '<root><parent1>11</parent1><parent2>12</parent2><parent3>13</parent3></root>';

var sura = '<?xml version=\'1.0\' ?><Farmer id="farmer_form"><start>2015-09-18T15:01:39.494+03</start>' +
    '<end>2015-09-18T15:39:11.618+03</end><today>2015-09-18</today><deviceid>358310068654699</deviceid>' +
    '</Farmer>'

describe("flatten json", function() {
    describe("default behavior", function() {
        it("should flatten a nested xml data and arrange sequential", function() {
            var jsonData = parser.toJson(xmlData, {
                object: true,
                coerce: true,
            });
            var results = xmlXlsTool.flatten(jsonData);
            expect(Object.keys(results).length).to.equal(5);
            expect(Object.keys(results)[0]).to.equal('child1');
            expect(Object.keys(results)[1]).to.equal('child11');
            expect(Object.keys(results)[2]).to.equal('child1111');
            expect(Object.keys(results)[3]).to.equal('child2');
            expect(Object.keys(results)[4]).to.equal('child3');
            expect(results).to.have.all.keys('child1', 'child2', 'child3', 'child11', 'child1111');
            expect(results.child1).to.have.a.property('child12', 12);
            expect(results.child2).to.have.a.property('child21', 20);
            expect(results.child3).to.have.a.property('child11', 101);
        });

        it('should flatten flat xml data', function() {
            var jsonData = parser.toJson(xmlFlatData, {
                object: true,
                coerce: true,
            });
            var results = xmlXlsTool.flatten(jsonData);
            expect(Object.keys(results).length).to.be.equal(3);
        });

        it("should not crash on empty objects", function() {
            var emptyXml = "";
            var jsonData = parser.toJson(emptyXml, {
                object: true
            });
            var results = xmlXlsTool.flatten(jsonData);
            expect(Object.keys(results).length).to.equal(0);
        });

        it("should not serialize functions", function() {
            var results = xmlXlsTool.flatten({
                value: 0,
                func: function() {
                    console.log("foo")
                }
            });
            expect(Object.keys(results).length).to.equal(1);
        });
        it("should honor filters", function() {
            var results = xmlXlsTool.flatten({
                value: 0,
                skip: 1,
                func: function() {
                    console.log("foo")
                }
            }, function(key, item) {
                if (key == "skip" || typeof(item) == 'function') return false;
                return true;
            });
            expect(Object.keys(results).length).to.equal(1);
        });
    });


});

describe.only('convert to excel', function() {
    it('should create workbook', function(done) {
        xmlXlsTool.buildXls(xmlData, function(xls) {
            expect(xls).to.be.null;
            done()
        })
    });

    it('should create excel file in filepath specified', function(done) {
        xmlXlsTool.buildXls(xmlData, {
            filepath: './',
            filename: 'output.xlsx'
        }, function(xls) {
            expect(xls).to.be.null;
            done()
        })
    });

    it('should create excel and return its buffer equivalent', function(done) {
        xmlXlsTool.buildXls(xmlData, {buffer:true}, function(xls) {
            expect(xls).not.to.be.null;
            done()
        })
    });
});
