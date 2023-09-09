let worksheet;
let currentRow = 1;
let currentCol = 0;
let lastRow;

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
        lastRow = range.e.r + 1;
        currentRow = getRandomInt(0, lastRow);
        currentCol = 0;
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
  const cols = ['A', 'B', 'C'];
  const cell = worksheet[`${cols[currentCol]}${currentRow}`];
  
  if (cell) {
    if (currentCol === 0) {
      characters.textContent = cell.v;
      pinyin.textContent = "";
      english.textContent = "";
    } else if (currentCol === 1) {
      pinyin.textContent = cell.v;
    } else if (currentCol === 2) {
      english.textContent = cell.v;
    }
  } else {
    console.log('Cell not found');
  }
}

function getRandomInt(min, max) {
  min = Math.ceil(min);
  max = Math.floor(max);
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

document.addEventListener('keydown', function(event) {
    if (event.code === 'Space') {
      currentCol += 1;
      
      if (currentCol > 2) {
        currentCol = 0;
        currentRow = getRandomInt(1, lastRow);
      }
      
      updateDisplay();
    }
  });
  