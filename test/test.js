var assert = require("assert"),
    S2C = require('../'),
    sheet = 'Test',
    file = 'test/test.xls';

beforeEach(function () {
    S2C.loadSheet(file);
});

describe('Static test (without setting values)', function () {
    describe('COS() with single value', function () {
        it('B8 should be -1', function () {
            assert.equal(-1, S2C.getValue(sheet, 'B8'));
        });
    });
    describe('COS() with intersect', function () {
        it('A6 should be -1', function () {
            assert.equal(-1, S2C.getValue(sheet, 'A6'));
        });
        it('B6 should be 1', function () {
            assert.equal(1, S2C.getValue(sheet, 'B6'));
        });
        it('C6 should be -1', function () {
            assert.equal(-1, S2C.getValue(sheet, 'C6'));
        });
    });
    describe('SUM()', function () {
        it('A7 should be 72', function () {
            assert.equal(72, S2C.getValue(sheet, 'A7'));
        });
    });
    describe('SUMIF()', function () {
        it('B7 should be 9', function () {
            assert.equal(9, S2C.getValue(sheet, 'B7'));
        });
    });
    describe('POWER() with single value', function () {
        it('C8 should be 1024', function () {
            assert.equal(1024, S2C.getValue(sheet, 'C8'));
        });
    });
    describe('POWER() with intersect', function () {
        it('C7 should be 10077696', function () {
            assert.equal(10077696, S2C.getValue(sheet, 'C7'));
        });
    });
    describe('OFFSET()', function () {
        it('A8 should be 7', function () {
            assert.equal(7, S2C.getValue(sheet, 'A8'));
        });
    });
});