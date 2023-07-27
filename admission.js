angular.module('myApp', [])
  .controller('AdmissionController', function ($scope) {
    $scope.formData = {};

    $scope.submitForm = function () {
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
      var row = Object.keys($scope.formData).map(function (key) {
        return $scope.formData[key];
      });
      XLSX.utils.sheet_add_aoa(worksheet, [row], { origin: -1 });

      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Contact Form Data');

      // Convert the workbook to a binary data object
      var wbout = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
      var blob = new Blob([wbout], { type: 'application/octet-stream' });

      // Save the Excel file using FileSaver.js
      saveAs(blob, 'contact_form_data.xlsx');

      // Show success message
      $scope.successMessage = 'Form submitted successfully!';
      $scope.showSuccessMessage = true;

      // Clear the form after submission
      $scope.formData = {};
    };
  });
