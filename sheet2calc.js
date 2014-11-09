/*
 * Description: Excel Calculator.
 * Limitations:
 *      - only JavaScript operators are allowed (e.g. ^-operator is not supported, use POWER() instead)
 *      - only alphanumeric sheet names are allowed (e.g. white spaces or dashes are not supported)
 *      - only functions provided by formula.js are supported (e.g. lots of financial formulas are not implemented)
 *      - special formula formats like RC mode are not supported
 *      - cell references like OFFSET are not supported
 *      - implicit and explicit cell unions and intersections are not supported
 *      - a massive security issue is the use of eval(): spreadsheet content can inject JS code
 *      - ... probably lots of other issues
 */

var FileParser = require('j'),
    FormulaParser = require('excel-formula'),
    FormulaFunctions = require('formulajs'),
    _ = require('underscore'),
    inputValues = {},
    cache = {}, // value cache
    workbook = {};

var loadSheet = function (filename, skipTest, debug) {
    emptyCache();
    workbook = FileParser.readFile(filename)[1];
    if (skipTest !== true) {
        testWorkbook(workbook, debug);
    }
};

var setValues = function (values) {
    emptyCache();
    inputValues = values;
};

var setValue = function (sheet, cell, value) {
    inputValues[sheet + '!' + cell] = value;
};

var getValue = function (sheet, cell) {
    // test input values
    if (inputValues[sheet + '!' + cell]) {
        return inputValues[sheet + '!' + cell];
    }

    // test local cache
    if (cache[sheet + '!' + cell]) {
        return cache[sheet + '!' + cell];
    }

    // the cell has a static value
    var formula = workbook.Sheets[sheet][cell].f;
    if (!formula) {
        var cellValue = workbook.Sheets[sheet][cell].v;
        return cache[sheet + '!' + cell] = _.isFinite(cellValue) ? parseFloat(cellValue) : cellValue;
    }

    // the cell has a formula
    return cache[sheet + '!' + cell] = calculateFormula(formula, sheet);
};

var getValueByReference = function (reference) {
    var sheet = reference.split("!")[0],
        cell = reference.split("!")[1];
    return getValue(sheet, cell);
};

/* private */
var testWorkbook = function (workbook, fullLog) {
    var tmpInputValues = inputValues,
        errors = [],
        numberOfTests = 0;

    // empty set values
    setValues({});

    // iterates over all cells
    _.forEach(workbook.Sheets, function (sheet, sheetName) {
        _.forEach(sheet, function (cell, cellName) {
            if (!_.isUndefined(cell.v)) {
                numberOfTests++;
                try {
                    var result = getValue(sheetName, cellName);
                    // test calculation against the cell value
                    if (!_.isEqual(String(result), String(cell.v))) {
                        errors[errors.length] = formatErrorMessage(sheetName, cellName, cell.f, result, cell.v);
                    }
                } catch (e) {
                    errors[errors.length] = formatErrorMessage(sheetName, cellName, cell.f, e, cell.v);
                }
            }
        });
    });

    if (!_.isEmpty(errors)) {
        throw new Error("Calculation of " + errors.length + "/" + numberOfTests + " (" +
        Math.round(errors.length * 100 / numberOfTests) +
        "%) cells has failed" + (fullLog === true ? ": \n" + errors.join('\n') : ""));
    }

    // re-adding set values
    setValues(tmpInputValues);
};

/* private */
var formatErrorMessage = function (sheet, cell, formula, result, cellValue) {
    return "Calculation failed for '" + sheet + "!" + cell + "' (" + formula + "): the calculated value '" + result
        + "' is not equal to the stored value '" + cellValue + "'";
};

/* private */
var calculateFormula = function (formula, sheet) {
    // convert formula to js expression
    formula = FormulaParser.toJavaScript(formula);

    // replacing references to other sheets
    formula = formula.replace(
        /([\w\u00E4\u00F6\u00FC\u00C4\u00D6\u00DC\u00df]+![A-Z]+[0-9]+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "getValueByReference(\"$1\")");

    // replacing local (in the current sheet) references
    formula = formula.replace(/([A-Z]+[0-9]+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "getValue(\"" + sheet + "\", \"" + "$1\")");

    // replacing function calls
    formula = formula.replace(/([A-Z][A-Z]+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "FormulaFunctions.$1");

    // considering %-signs, it's a hack, as well
    formula = formula.replace(/((\d+.){0,1}\d+%+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "($1 / 100)");
    formula = formula.replace(/(%+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g, "");

    // evaluate expression, quick and dirty, better would be to use the ATS directly
    var evaluated = eval(formula);

    // convert to number if possible
    return _.isFinite(evaluated) ? parseFloat(evaluated) : evaluated;
};

/* private */
var emptyCache = function () {
    cache = {};
};

var S2C = {
    loadSheet: loadSheet,
    setValues: setValues,
    setValue: setValue,
    getValue: getValue,
    getValueByReference: getValueByReference,
    version: "0.0.2"
};

if (typeof module !== 'undefined') module.exports = S2C;
