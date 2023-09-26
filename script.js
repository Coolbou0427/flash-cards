let worksheet;
let currentRow = 1;
let currentCol = 0;
let lastRow;
let randomRows = [];
let randomIndex = 0;

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
        
        randomRows = [...Array(lastRow).keys()].slice(1); // Create an array [1, 2, ..., lastRow-1]
        shuffleArray(randomRows); // Shuffle the array to get random order
        randomIndex = 0; // Reset the index
        currentRow = randomRows[randomIndex]; // Get the first random row
        
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

function getRandomInt() {
  randomIndex += 1;
  
  if (randomIndex >= randomRows.length) {
    shuffleArray(randomRows); // Reshuffle the array when all rows have been used
    randomIndex = 0; // Reset the index
  }
  
  return randomRows[randomIndex];
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}

document.addEventListener('keydown', function(event) {
  if (event.code === 'Space') {
    currentCol += 1;
    
    if (currentCol > 2) {
      currentCol = 0;
      currentRow = getRandomInt(); // Get the next random row
    }
    
    updateDisplay();
  }
});
