document.addEventListener('DOMContentLoaded', () => {
  const container = document.getElementById('excel-grid');
  
  // Load your CSV file
  fetch1('C:\Users\Administrator\Desktop\WORK SHEET\~$DAILY STOCKSHEET.xlsx')
  fetch('C:\Users\Administrator\Desktop\WORK SHEET\MONTLYREVSTOCKSHEET.xlsx')
    .then(response => response.text())
    .then(csvData => {
      // Parse CSV to JSON
      const parsedData = Papa.parse(csvData, { header: false }).data;
      
      // Initialize Handsontable with your data
      const hot = new Handsontable(container, {
        data: parsedData,
        colHeaders: true,
        rowHeaders: true,
        formulas: true,
        cells(row, col) {
          const cellProperties = {};
          // Lock columns with formulas (e.g., column 3 = "Total")
          if (col === 3) { // Change this to your formula column index
            cellProperties.readOnly = true;
          }
          return cellProperties;
        }
      });

      // Block formula edits and send approval requests
      hot.addHook('beforeChange', (changes) => {
        const [row, col, oldValue, newValue] = changes[0];
        if (col === 3) { // If editing a formula column
          sendApprovalRequest(row, col, newValue);
          alert("Formula changes require owner approval!");
          return false; // Block the edit
        }
      });
    });

  // Send approval request to Firebase
  function sendApprovalRequest(row, col, newValue) {
    db.collection("approvals").add({
      row: row,
      col: col,
      newValue: newValue,
      timestamp: firebase.firestore.FieldValue.serverTimestamp(),
      status: "pending"
    }).then(() => {
      document.getElementById('approval-alert').style.display = 'block';
    });
  }
});