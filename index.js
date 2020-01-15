'use strict';
const excelToJson = require('convert-excel-to-json');
const Excel = require('exceljs');

const headings = excelToJson({
	sourceFile: 'data/data.xlsx',
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

const result = excelToJson({
	sourceFile: 'data/data.xlsx',
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

/* Retrieve the heading row from JSON data */
const getHeadings = (json) => {
	let list = [];
	for (let i = 0; i < 4; i++) {
		json[i]['DG'] = (json[i]['DG'] * 100).toFixed(2) + ' %';
		json[i]['DR'] = (json[i]['DR'] * 100).toFixed(2) + ' %';
		json[i]['DW'] = (json[i]['DW'] * 100).toFixed(2) + ' %';
		json[i]['DZ'] = (json[i]['DZ'] * 100).toFixed(2) + ' %';
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
			list.push(record);
			continue;
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
const generateExcel = (json) => {
	const options = {
		filename: 'report.xlsx',
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
		console.log('excel file cretaed');
	});
};

const records = findRecord('Nb Network Solutions', result.DLAR);
const headers = getHeadings(headings.DLAR);

generateExcel([ ...headers, ...records ]);
