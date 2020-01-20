console.log('App loading...');
console.log('JS script "query.js" loaded...');

let headerRows;
let resultRows;
let selected = [ [], [] ];
let file_path;
let data = [];

$(document).ready(() => {
	console.log('Modules Loaded. App main screen displayed...');

	/* After submitting the file_path, run the script and load the company/employee list */
	$('#upload').click((event) => {
		console.log('User uploaded file: ' + document.getElementById('filename').files[0]);
		console.log('Uploaded File path: ' + document.getElementById('filename').files[0].path);
		file_path = document.getElementById('filename').files[0].path;
		console.log('Retrieving Employee Names...\n');
		headerRows = headings(file_path);
		resultRows = result(file_path);
		const employees = listEmployees(resultRows.DLAR);
		for (let i = 0; i < employees.length; i++) {
			$('#employees').append(new Option(employees[i], employees[i]));
		}
		console.log('Employee names loaded...');
	});

	/* Every time the user selects a company, add it to the array */
	$('#companies').on('change', function(e) {
		var optionSelected = $('option:selected', this);
		if (!selected[1].includes(this.value)) {
			selected[1].push(this.value);
			console.log('Companies selected: ' + selected[1]);
			$('#co-list').after('<p>' + this.value + '</p>');
		}
	});

	/* Every time the user selects a employee, add it to the array */
	$('#employees').on('change', function(e) {
		var optionSelected = $('option:selected', this);
		if (!selected[0].includes(this.value)) {
			selected[0].push(this.value);
			console.log('Employees selected: ' + selected[0]);
			$('#emp-list').after('<p>' + this.value + '</p>');
		}
	});

	/* When Employee list is confirmed, generate the filtered company list */
	$('#emp-btn').click((event) => {
		const companies = listCompanies(selected[0], resultRows.DLAR).sort();
		for (let i = 0; i < companies.length; i++) {
			$('#companies').append(new Option(companies[i], companies[i]));
		}
		console.log('Company list loaded...');
	});

	/* When Company List is confirmed, run a query for each item on the list */
	$('#cp-btn').click((event) => {
		console.log('Generating report document based on the query data...');
		const headers = getHeadings(headerRows.DLAR);
		data = [ ...headers ];
		for (let i = 0; i < selected[1].length; i++) {
			data = [ ...data, ...findRecord(selected[0], selected[1][i], resultRows.DLAR) ];
		}
		console.log('Finshed query...');
	});

	/* Generate the report after user clicks the 'generate' button */
	$('#generate').click((event) => {
		console.log('Generating report...');
		if (file_path && data) {
			const new_path = parse_path(file_path);
			generateExcel(new_path, data);
			console.log('Report generated....');
		} else {
			console.log('No file_path or data generated. Cant generate report');
		}
	});
});
