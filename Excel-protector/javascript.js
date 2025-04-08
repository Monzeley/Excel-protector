let hot; // Stores the Handsontable instance

// Load CSV file into the grid
function loadCSV(filename) {
  fetch(filename)
    .then(response => response.text())
    .then(csvData => {
      const parsedData = Papa.parse(csvData, { header: false }).data;
      
      // Clear previous grid if it exists
      const container = document.getElementById('excel-grid');
      if (hot) hot.destroy();
      
      // Create new grid
      hot = new Handsontable(container, {
        data: parsedData,
        colHeaders: true,
        rowHeaders: true,
        formulas: true,
        cells(row, col) {
          const cellProperties = {};
          // Lock column 3 (Total) - change this to your formula column
          if (col === 3) cellProperties.readOnly = true;
          return cellProperties;
        }
      });

      // Block formula edits
      hot.addHook('beforeChange', (changes) => {
        const [row, col, oldValue, newValue] = changes[0];
        if (col === 3) { // If editing formula column
          sendApprovalRequest(row, col, newValue);
          alert("Owner must approve formula changes!");
          return false; // Block the edit
        }
      });
    });
}

// Send approval request to Firebase
function sendApprovalRequest(row, col, newValue) {
  const db = firebase.firestore();
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