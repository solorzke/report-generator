'use strict';
const excelToJson = require('convert-excel-to-json');
const Excel = require('exceljs');

console.log("Node.js script: 'index.js' loaded...");

/* Retrieve the root path of the file that was submitted and return a new file path */
const parse_path = (file_path) => {
	const index = file_path.lastIndexOf('/');
	const path = file_path.substr(0, index + 1);
	return path + 'report.xlsx';
};

const headings = (file_path) => {
	return excelToJson({
		sourceFile: file_path,
		sheets: [ 'DLAR' ],
		columnToKey: {
			B: 'B',
			L: 'L',
			N: 'N',
			O: 'O',
			P: 'P',
			AI: 'AI',
			AJ: 'AJ',
			AK: 'AK',
			AL: 'AL',
			DG: 'DG',
			DR: 'DR',
			DW: 'DW',
			DY: 'DY',
			DZ: 'DZ'
		}
	});
};

const result = (file_path) => {
	return excelToJson({
		sourceFile: file_path,
		rows: 5,
		sheets: [ 'DLAR' ],
		columnToKey: {
			L: 'L',
			N: 'N',
			O: 'O',
			P: 'P',
			AI: 'AI',
			AJ: 'AJ',
			AK: 'AK',
			AL: 'AL',
			DG: 'DG',
			DR: 'DR',
			DW: 'DW',
			DY: 'DY',
			DZ: 'DZ'
		}
	});
};

// const headings = excelToJson({
// 	sourceFile: 'data/data.xlsx',
// 	sheets: [ 'DLAR' ],
// 	columnToKey: {
// 		B: 'B',
// 		L: 'L',
// 		N: 'N',
// 		O: 'O',
// 		P: 'P',
// 		AI: 'AI',
// 		AJ: 'AJ',
// 		AK: 'AK',
// 		AL: 'AL',
// 		DG: 'DG',
// 		DR: 'DR',
// 		DW: 'DW',
// 		DY: 'DY',
// 		DZ: 'DZ'
// 	}
// });

// const result = excelToJson({
// 	sourceFile: 'data/data.xlsx',
// 	rows: 5,
// 	sheets: [ 'DLAR' ],
// 	columnToKey: {
// 		L: 'L',
// 		N: 'N',
// 		O: 'O',
// 		P: 'P',
// 		AI: 'AI',
// 		AJ: 'AJ',
// 		AK: 'AK',
// 		AL: 'AL',
// 		DG: 'DG',
// 		DR: 'DR',
// 		DW: 'DW',
// 		DY: 'DY',
// 		DZ: 'DZ'
// 	}
// });

/* Retrieve the heading row from JSON data */
const getHeadings = (json) => {
	let list = [];
	for (let i = 0; i < 4; i++) {
		if (i != 0 && i != 3) {
			json[i]['DG'] = (json[i]['DG'] * 100).toFixed(2) + ' %';
			json[i]['DR'] = (json[i]['DR'] * 100).toFixed(2) + ' %';
			json[i]['DW'] = (json[i]['DW'] * 100).toFixed(2) + ' %';
			json[i]['DZ'] = (json[i]['DZ'] * 100).toFixed(2) + ' %';
		}
		list.push(json[i]);
	}
	delete list[3]['B'];
	return list;
};

/* Return a array of company names from the JSON data */
const listCompanies = (json) => {
	let list = [];
	for (let i = 0; i < json.length; i++) {
		if (json[i].hasOwnProperty('L')) {
			let record = JSON.stringify(json[i]['L']).trim();
			record = record.slice(1, record.length - 1);
			if (!list.includes(record)) {
				list.push(record);
			} else {
				continue;
			}
		}
	}
	return list;
};

/* Find the record of the company name via JSON. Return as obj */
const findRecord = (companyName, json) => {
	let data = [];

	for (let i = 0; i < json.length; i++) {
		if (json[i].hasOwnProperty('L')) {
			let record = JSON.stringify(json[i]['L']).trim();
			//Remove double quotes surrounding the name
			record = record.slice(1, record.length - 1);
			if (record === companyName) {
				json[i]['DG'] = (json[i]['DG'] * 100).toFixed(2) + ' %';
				json[i]['DR'] = (json[i]['DR'] * 100).toFixed(2) + ' %';
				json[i]['DW'] = (json[i]['DW'] * 100).toFixed(2) + ' %';
				json[i]['DZ'] = (json[i]['DZ'] * 100).toFixed(2) + ' %';
				data.push(json[i]);
			}
		}
	}
	const result = data.length !== 0 ? data : 'Company Name does not exist!';
	return result;
};

/* Generate the XLSX file based on the sorted data */
const generateExcel = (file_path, json) => {
	const options = {
		filename: file_path,
		useStyles: true,
		useSharedStrings: true
	};

	const workbook = new Excel.stream.xlsx.WorkbookWriter(options);
	const worksheet = workbook.addWorksheet('DLAR');
	worksheet.columns = [
		{ header: '', key: 'B' },
		{ header: '', key: 'L' },
		{ header: '', key: 'N' },
		{ header: '', key: 'O' },
		{ header: '', key: 'P' },
		{ header: '', key: 'AI' },
		{ header: '', key: 'AJ' },
		{ header: '', key: 'AK' },
		{ header: '', key: 'AL' },
		{ header: '', key: 'DG' },
		{ header: '', key: 'DR' },
		{ header: '', key: 'DW' },
		{ header: '', key: 'DY' },
		{ header: '', key: 'DZ' }
	];

	for (let i = 0; i < json.length; i++) {
		worksheet.addRow(json[i]).commit();
	}

	workbook.commit().then(() => {
		console.log('File created. Stored in ' + file_path);
	});
};

// const r = result('/Users/solorzke/Downloads/prepaid_daily_pulse_naws(1).xlsx').DLAR;
// const records = findRecord('Nb Network Solutions', r);
// const h = headings('/Users/solorzke/Downloads/prepaid_daily_pulse_naws(1).xlsx');
// // //const headers = getHeadings(h);
// // console.log(records);
// // //generateExcel([ ...headers, ...records ]);
// const headers = getHeadings(h.DLAR);
// console.log(headers);
