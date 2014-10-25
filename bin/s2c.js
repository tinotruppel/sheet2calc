var s2c;
try { s2c = require('../'); } catch(e) { s2c = require('sheet2calc'); }
var fs = require('fs'), program = require('commander');
program
	.version(s2c.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified file (- for stdin)')
	.option('-i, --input <input>', 'a set of cells and its values')
	.option('-o, --output <output>', 'a set of cells to calculate')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-N, --sheet-index <idx>', 'use specified sheet index (0-based)')
	.option('-l, --list-sheets', 'list sheet names and exit');

program.on('--help', function() {
	console.log('  Takes a set of cells as input and calculates the values of an other set of cells.');
	console.log('  Output format is JSON.');
	console.log('  Support email: tino.truppel@gmail.com');
});

program.parse(process.argv);

// TODO
