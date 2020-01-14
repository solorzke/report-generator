const Excel = require('exceljs');

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

var data;

for (let i = 1; i <= 10; i++) {
	data = {
		id: i,
		'first name': 'name ' + i,
		ph: '012014520' + i
	};

	worksheet.addRow(data).commit();
}

workbook.commit().then(function() {
	console.log('excel file cretaed');
});
