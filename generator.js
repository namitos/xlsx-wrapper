var fs = require('fs');
var os = require('os');

var xlsx = require('xlsx-style');

function datenum(v, date1904) {
	if (date1904) v += 1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function convertData(data, opts) {
	var ws = {};
	var range = {
		s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}
	};
	for (var R = 0; R != data.length; ++R) {
		for (var C = 0; C != data[R].length; ++C) {
			if (range.s.r > R) range.s.r = R;
			if (range.s.c > C) range.s.c = C;
			if (range.e.r < R) range.e.r = R;
			if (range.e.c < C) range.e.c = C;
			var cell = {
				v: data[R][C]
			};
			if (cell.v == null) continue;
			var cell_ref = xlsx.utils.encode_cell({
				c: C,
				r: R
			});

			if (typeof cell.v === 'number') {
				cell.t = 'n';
			} else if (typeof cell.v === 'boolean') {
				cell.t = 'b';
			} else if (cell.v instanceof Date) {
				cell.t = 'n';
				cell.z = xlsx.SSF._table[14];
				cell.v = datenum(cell.v);
			} else {
				cell.t = 's';
			}

			var borderStyle = {
				style: 'thin',
				color: {
					rgb: "44444400"
				}
			};
			cell.s = {
				border: {
					top: borderStyle,
					right: borderStyle,
					bottom: borderStyle,
					left: borderStyle
				}
			};

			ws[cell_ref] = cell;
		}
	}
	if (range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range);
	return ws;
}

process.on('message', function (message) {
	var sheets = {};
	Object.keys(message.sheets).forEach(function (sheetName) {
		sheets[sheetName] = convertData(message.sheets[sheetName]);
	});

	var path = os.tmpdir() + '/' + new Date().valueOf() + '.xlsx';
	xlsx.writeFile({
		SheetNames: Object.keys(sheets),
		Sheets: sheets
	}, path);
	process.send({
		result: path
	});

});
