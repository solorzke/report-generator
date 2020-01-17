console.log('App loading...');
console.log('JS script "query.js" loaded...');

let headerRows;
let resultRows;
let selected = [];
let file_path;

$(document).ready(() => {
	console.log('Modules Loaded. App main screen displayed...');

	/* After submitting the file_path, run the script and load the company list */
	$('#upload').click((event) => {
		console.log('User uploaded file: ' + document.getElementById('filename').files[0]);
		console.log('Uploaded File path: ' + document.getElementById('filename').files[0].path);
		file_path = document.getElementById('filename').files[0].path;
		console.log('Retrieving Company Names...\n');
		headerRows = headings(file_path);
		resultRows = result(file_path);
		const companies = listCompanies(resultRows.DLAR);
		for (let i = 0; i < companies.length; i++) {
			$('#companies').append(new Option(companies[i], companies[i]));
		}
		console.log('Company names loaded...');
	});

	/* Every time the user selects a company, add it to the array */
	$('select').on('change', function(e) {
		var optionSelected = $('option:selected', this);
		if (!selected.includes(this.value)) {
			selected.push(this.value);
			console.log('Companies selected: ' + selected);
			$('#co-list').after('<p>' + this.value + '</p>');
		}
	});

	/* Generate the report after user clicks the 'generate' button */
	$('#generate').click((event) => {
		console.log('Generating report document based on the query data...');
		const headers = getHeadings(headerRows.DLAR);
		let data = [ ...headers ];
		for (let i = 0; i < selected.length; i++) {
			data = [ ...data, ...findRecord(selected[i], resultRows.DLAR) ];
		}
		console.log('Finshed query. Generating report...');
		const new_path = parse_path(file_path);
		generateExcel(new_path, data);
		console.log('Report generated....');
	});
});
