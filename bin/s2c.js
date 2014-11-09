var S2C;
try {
    S2C = require('../');
} catch (e) {
    S2C = require('sheet2calc');
}
var program = require('commander');
program
    .version(S2C.version)
    .usage('[options] <file> <sheet> <cell>')
    .option('-i, --input <input>', 'a set of cells and its values')
    .option('-s, --skipTest', 'skips the initial calculation tests of the sheet')
    .option('-d, --debug', 'debug');

program.on('--help', function () {
    console.log('  Takes a set of cells as input and calculates the value of an other cells.');
});

program.parse(process.argv);

var filename, sheet, cell = '', input = {}, skipTest, debug = false;
if (program.args[0]) filename = program.args[0];
if (program.args[1]) sheet = program.args[1];
if (program.args[2]) cell = program.args[2];
if (program.input) input = JSON.parse(program.input);
if (program.skipTest) skipTest = true;
if (program.debug) debug = true;

if (!filename || !sheet || !cell) {
    console.error("s2c: must specify a filename, sheet and a cell");
    process.exit(1);
}

S2C.loadSheet(filename, skipTest, debug);
S2C.setValues(input);
console.log(S2C.getValue(sheet, cell));