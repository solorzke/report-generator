'use strict';
const excelToJson = require('convert-excel-to-json');
const Excel = require('exceljs');

const result = excelToJson({
	sourceFile: 'data/data.xlsx',
	rows: 5,
	sheets: [ 'DLAR' ],
	columnToKey: {
		L: 'Company Name',
		N: 'Address',
		O: 'City',
		P: 'State',
		AI: 'Jan-19',
		AJ: 'Nov-19',
		AK: 'Dec-19',
		AL: 'Jan-20',
		DG: 'Comp 3rd Month (Current Month)',
		DR: 'Ports in Percent',
		DW: 'Prior Month Conversion Rate',
		DY: 'RPL Count',
		DZ: 'Conversion Rate'
	}
});

/* Return a array of company names from the JSON data */
const listCompanies = (json) => {
	let list = [];
	for (let i = 0; i < json.length; i++) {
		if (json[i].hasOwnProperty('Company Name')) {
			let record = JSON.stringify(json[i]['Company Name']).trim();
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
		if (json[i].hasOwnProperty('Company Name')) {
			let record = JSON.stringify(json[i]['Company Name']).trim();
			//Remove double quotes surrounding the name
			record = record.slice(1, record.length - 1);
			if (record === companyName) {
				data.push(json[i]);
			}
		}
	}
	const result = data.length !== 0 ? data : 'Company Name does not exist!';
	return result;
};

// console.log(findRecord('Nb Network Solutions', result.DLAR));
// console.log(setHeadings(findRecord('Company Name', result.DLAR))['Projection']);
// console.log(listCompanies(result.DLAR));

const generateExcel = (json) => {
	const options = {
		filename: 'report.xlsx',
		useStyles: true,
		useSharedStrings: true
	};

	const workbook = new Excel.stream.xlsx.WorkbookWriter(options);
	const worksheet = workbook.addWorksheet('DLAR');

	worksheet.columns = [
		{ header: 'Company Name', key: 'Company Name' },
		{ header: 'Address', key: 'Address' },
		{ header: 'City', key: 'City' },
		{ header: 'State', key: 'State' },
		{ header: 'Jan-19', key: 'Jan-19' },
		{ header: 'Nov-19', key: 'Nov-19' },
		{ header: 'Dec-19', key: 'Dec-19' },
		{ header: 'Jan-20', key: 'Jan-20' },
		{ header: 'Comp 3rd Month (Current Month)', key: 'Comp 3rd Month (Current Month)' },
		{ header: 'Ports in Percent', key: 'Ports in Percent' },
		{ header: 'Prior Month Conversion Rate', key: 'Prior Month Conversion Rate' },
		{ header: 'RPL Count', key: 'RPL Count' },
		{ header: 'Conversion Rate', key: 'Conversion Rate' }
	];

	for (let i = 0; i < json.length; i++) {
		worksheet.addRow(json[i]).commit();
	}

	workbook.commit().then(() => {
		console.log('excel file cretaed');
	});
};

const records = findRecord('Nb Network Solutions', result.DLAR);
generateExcel(records);
