//USAGE: node index.js  pathToCountryFile outputPath
console.warn(
  "USAGE: node index.js /pathToCountryFile /outputPath"
);
console.log("Please include file extensions");
console.log("Input is xlsx format and output is json");

const fs = require("fs");
const excelToJson = require("convert-excel-to-json");

const currentExcelFile = process.argv[2];
const outputFile = process.argv[3];

const result = excelToJson({
  sourceFile: currentExcelFile,
  header: {
    rows: 1,
  },
  columnToKey: {
    A: "value",
    B: "key",
  },
});

const getRowData = (rowValue) => {
  let updatedRowValue = rowValue.toLowerCase().split(' ').map((word, index) => {
    let wordToReturn = word
    // Keep these words as lowercase as long as they are not the first word
    if (['and', 'in', 'of', 'at', 'the'].includes(word) && index !== 0) {
      return word
    }
    //Uppercase the second character if first character is ( or -
    if (wordToReturn.charAt(0) === '(' || wordToReturn.charAt(0) === ',') {
      wordToReturn = wordToReturn.charAt(0) + wordToReturn.charAt(1).toUpperCase() + wordToReturn.slice(2)
    }
    // If word contains a - then uppercase both sides of the -
    if (wordToReturn.includes('-')) {
      const splitWords = wordToReturn.split('-')
      wordToReturn = splitWords.map(singleWord => singleWord.charAt(0).toUpperCase() + singleWord.slice(1)).join('-')
    }
    // Uppercase the first character
    wordToReturn = wordToReturn.charAt(0).toUpperCase() + wordToReturn.slice(1)
    return wordToReturn
  }).join(' ')

  //Handle special case for the.  E.g Bahamas, the. Should be Bahanas, the
  if (updatedRowValue.includes(', the')) {
    console.log('updating ' + updatedRowValue)
    updatedRowValue = updatedRowValue.replace(', the', ', The')
    console.log('to ' + updatedRowValue)
  }
  return updatedRowValue
}

const editedResult = result.Sheet1.map(row => {
  return {
    value: getRowData(row.value),
    key: row.key.toString()
  }
}).filter(row => row.key !== '-' && row.key !== '')

fs.writeFileSync(outputFile, JSON.stringify(editedResult), "utf8");

console.log("DONE");
