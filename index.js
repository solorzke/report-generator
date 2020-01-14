'use strict';
const excelToJson = require('convert-excel-to-json');
const xlsx = require('json-as-xlsx');

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

const records = findRecord('Nb Network Solutions', result.DLAR);
const columns = [
	{ label: 'Company Name', value: 'company name' },
	{ label: 'Address', value: 'address' },
	{ label: 'City', value: 'city' },
	{ label: 'State', value: 'state' },
	{ label: 'Jan-19', value: 'jan-19' },
	{ label: 'Nov-19', value: 'nov-19' },
	{ label: 'Dec-19', value: 'dec-19' },
	{ label: 'Jan-20', value: 'jan-20' },
	{ label: 'Comp 3rd Month (Current Month)', value: 'comp 3rd month (current month)' },
	{ label: 'Ports in Percent', value: 'ports in percent' },
	{ label: 'Prior Month Conversion Rate', value: 'prior month conversion rate' },
	{ label: 'RPL Count', value: 'rpl count' },
	{ label: 'Conversion Rate', value: 'conversion rate' }
];

let content = [ ...records ];

const settings = {
	sheetName: 'DLAR',
	fileName: 'Report'
};

xlsx(columns, content, settings);
