$(document).ready(function () {
  $('#admissionForm').submit(function (event) {
    event.preventDefault();

    // Create an Excel workbook
    var workbook = XLSX.utils.book_new();
    // Create a new worksheet
    var worksheet = XLSX.utils.aoa_to_sheet([
      [
        'Name',
        'DOB',
        'Gender',
        'Community',
        '10th Regno',
        '10th %',
        '12th Regno',
        '12th %',
        'Mobileno',
        'Email',
        "Father's Name",
        "Father's Occupation",
        'Annual income',
        'Aadhar no',
        'Address',
        'Course',
      ],
    ]);

    // Add the form data to the worksheet
    var data = $(this).serializeArray();
    var row = data.map((item) => Object.values(item).map((value) => value));
    XLSX.utils.sheet_add_aoa(worksheet, row, { origin: -1 });

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Contact Form Data');

    // Convert the workbook to a binary data object
    var wbout = XLSX.write(workbook, { type: 'array', bookType: 'xls' });
    var blob = new Blob([wbout], { type: 'application/octet-stream' });

    // Save the Excel file using FileSaver.js
    saveAs(blob, 'details.xls');

    // Show success message
    $("#successMessage").text("Form submitted successfully!");

    // Clear the form after submission
    $(this)[0].reset();
  });
});
