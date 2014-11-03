/*!
 * Description: Excel Calculator.
 * Limitations:
 *      - only JavaScript operators are allowed (e.g. ^-operator is not supported, use POWER() instead)
 *      - only alphanumeric sheet names are allowed (e.g. white spaces or dashes are not supported)
 *      - only functions provided by formula.js are supported
 *      - ... lots of other issues
 */

var FileParser = require('j'),
    FormulaParser = require('excel-formula'),
    FormulaFunctions = require('formulajs'),
    inputValues = {},
    cache = {}, // value cache, used per calculation
    workbook = {};

var loadSheet = function (uri) {
    emptyCache();
    workbook = FileParser.readFile(uri)[1];
};

var setValues = function (values) {
    emptyCache();
    inputValues = values;
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
        cache[sheet + '!' + cell] = isNumber(cellValue) ? parseFloat(cellValue) : cellValue;
        return cache[sheet + '!' + cell];
    }

    // the cell has a formula, convert formula to js expression
    formula = FormulaParser.toJavaScript(formula);

    // replacing references to other sheets
    formula = formula.replace(/([\w\u00E4\u00F6\u00FC\u00C4\u00D6\u00DC\u00df]+![A-Z]+[0-9]+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "getValueByReference(\"$1\")");

    // replacing local (in the current sheet) references
    formula = formula.replace(/([A-Z]+[0-9]+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "getValue(\"" + sheet + "\", \"" + "$1\")");

    // replacing function calls
    formula = formula.replace(/([A-Z][A-Z]+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "FormulaFunctions.$1");

    // considering %-signs
    formula = formula.replace(/((\d+.){0,1}\d+%+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g,
        "($1 / 100)");
    formula = formula.replace(/(%+)(?=([^"\\]*(\\.|"([^"\\]*\\.)*[^"\\]*"))*[^"]*$)/g, "");

    // evaluate expression, quick and dirty, better would be to use the ATS directly
    var evaluated = eval(formula);

    // convert to number if possible
    cache[sheet + '!' + cell] = isNumber(evaluated) ? parseFloat(evaluated) : evaluated;
    return cache[sheet + '!' + cell];

};

var getValueByReference = function (reference) {
    var sheet = reference.split("!")[0],
        cell = reference.split("!")[1];
    return getValue(sheet, cell);
};

/* private */
var isNumber = function (n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
};

/* private */
var emptyCache = function() {
    cache = {};
};

var S2C = {
    loadSheet: loadSheet,
    setValues: setValues,
    getValue: getValue,
    getValueByReference: getValueByReference,
    version: "0.0.1"
};

if (typeof module !== 'undefined') module.exports = S2C;
