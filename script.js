let worksheet;
let currentRow = 1;
let sequence = [];
let availableRows = [];

const spreadsheet = document.getElementById('upload').addEventListener('change', handleFileSelect);
const characters = document.getElementById('characters');
const pinyin = document.getElementById('pinyin');
const english = document.getElementById('english');
const error = document.getElementById('error');

function handleFileSelect(event) {
  const file = event.target.files[0];

  if (file) {
    console.log('File selected:', file.name);

    if (file.name.endsWith('.xlsx')) {
      document.getElementById('actual').style.display = 'none';
      document.getElementById('titles').style.display = 'none';
      error.textContent = "";
      const reader = new FileReader();

      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const lastRow = range.e.r + 1;

        availableRows = [...Array(lastRow).keys()].slice(1); // Create an array [1, 2, ..., lastRow-1]
        currentRow = getRandomInt(); // Get the first random row

        updateDisplay();
      };

      reader.readAsArrayBuffer(file);
    } else {
      error.textContent = 'The file is not a .xlsx file';
    }
  } else {
    error.textContent = 'No file selected';
  }
}

function updateDisplay() {
  // If sequence is empty, populate it based on random chance
  if (sequence.length === 0) {
    characters.innerHTML = "&nbsp;";
    pinyin.innerHTML = "&nbsp;";
    english.innerHTML = "&nbsp;";
    const randomChance = Math.random();
    sequence = randomChance < 0.33 ? ['C', 'A', 'B'] : ['A', 'B', 'C'];
  }

  const actualCol = sequence.shift(); // Remove and return the first element from sequence

  const cell = worksheet[`${actualCol}${currentRow}`];

  if (cell) {
    if (actualCol === 'A') {
      characters.textContent = cell.v;
    } else if (actualCol === 'B') {
      pinyin.textContent = cell.v;
    } else if (actualCol === 'C') {
      english.textContent = cell.v;
    }
  } else {
    console.log('Cell not found');
  }
}

function getRandomInt() {
  const randomIndex = Math.floor(Math.random() * availableRows.length);
  const randomRow = availableRows.splice(randomIndex, 1)[0];
  
  if (availableRows.length === 0) {
    console.log("Out of rows, restarting");
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const lastRow = range.e.r + 1;
    availableRows = [...Array(lastRow).keys()].slice(1);
  }
  
  return randomRow;
}

document.addEventListener('keydown', function(event) {
  if (event.code === 'Space') {
    if (sequence.length === 0) {
      currentRow = getRandomInt();
    }
    updateDisplay();
  }
});
