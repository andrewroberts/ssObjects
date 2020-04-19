// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO

// SsObjects.gs
// ============

// Test Sheet: https://docs.google.com/spreadsheets/d/1tAYu2STTAebhwq67LArlsVH5k1QmIMPmOxuqI7HYWrY/edit#gid=0

function isArray(value) {
  return (typeof value === 'object' && Object.prototype.toString.call(value) === '[object Array]')
}  

/**
 * @params {Sheet} sheet
 * @params {array} data - including the header
 */

function clearAndSet(sheet, data) {
  
  if (data.length < 1) {return}
  data.shift() // Remove the headers
  var numberOfRows = sheet.getLastRow()
  
  if (numberOfRows > 1) {
    sheet.getRange(2, 1, numberOfRows - 1, sheet.getLastColumn()).clearContent()
  }
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data)
  
} // clearAndSet()

/**
 * Convert table of data, with the headers in the first row
 * into an object, or add to an existing object. Note how if there are 
 * mulitple rows with the same ID the new values are stored in an array
 *
 *
 * [
 *   [id_header, header2,     header3, ...],
 *   [id1,       row1_value2, row1_value3, ...],
 *   [id2,       row2_value2, row2_value3, ...],
 *   [id1,       row3_value2, row3_value3, ...],
 * }
 * 
 * =>
 *
 * {
 *   id1: {
 *     [header2]: [row1_value2, row3_value2],
 *     [header3]: row1_value3
 *   }, 
 *   id2: {
 *     [header2]: row2_value2,
 *     [header3]: row2_value3
 *   },    
 * }
 *
 * @param {object} 
 *   {array} data - first row is headers
 *   {string} id - header used for id 
 *   {object} objects [OPTIONAL]
 *
 * @return {object} data converted to an object
 */

function addArrayToObject(config) {

  var data = config.data || (function() {throw new Error('No data')})()
  var idHeaderName = config.id || (function() {throw new Error('No ID header')})()
  var objects = config.objects || {}
    
  var headers = data[0]
  
  data.slice(1).forEach(function(row) {
    
    var id = null
    
    row.forEach(function(nextValue, index) {
      
      var nextHeader = headers[index].trim()
      
      if (nextHeader === idHeaderName) {
        
        if (!id) {
        
          id = nextValue
          
          if (objects[id] === undefined) {          
            objects[id] = {}                
          }
        }
        
      } else {

        if (!id) {
          throw new Error('No ID "' + idHeaderName + '" found for this object yet')
        }

        nextValue = nextValue || ''
        
        if (nextValue instanceof Date || isISODateString_(nextValue)) {
          nextValue = new Date(nextValue)
        }
                
        if (objects[id][nextHeader] === undefined) {

          objects[id][nextHeader] = nextValue
        
        } else {
          
          var existingValue = objects[id][nextHeader]
          
          if (isArray(existingValue)) {
          
            if (existingValue.length === 1) {throw new Error('Array field with only one value')}
          
            objects[id][nextHeader].push(nextValue)
            
          } else {          
          
            if (nextValue !== '') {
              objects[id][nextHeader] = [existingValue, nextValue]
            }
          }          
        }        
      }       
      
    }) // for each cell
    
  }) // for each row
  
  return objects  
  
} // SsObjects.addArrayToObject()

/**
 * Convert an object into a 2D table of data, with the headers in the first row,
 * or add to an existing table:
 *
 * {
 *   id1: {
 *     [header2]: row1_value2,
 *     [header3]: [row1_value3, row3_value3]
 *   }, 
 *   id2: {
 *     [header2]: row2_value2,
 *     [header3]: row2_value3
 *   },    
 * }
 *
 * => 
 *
 * [
 *   [id_header, header2,     header3, ...],
 *   [id1,       row1_value2, row1_value3, ...],
 *   [id1,       row1_value2, row3_value3, ...], 
 *   [id2,       row2_value2, row2_value3, ...],
 * }
 *
 * @param {object} 
 *   {string} id - header used for id 
 *   {object} objects
 *   {array}  data - first row is headers [OPTIONAL]
 *
 * @return {object} data converted to an object
 */

function addObjectsToArray(config) {

  var idHeaderName = config.id || (function() {throw new Error('No ID header')})()
  var objects = config.objects || (function() {throw new Error('No objects')})()
  var data = config.data || []

  var headers = getHeaders()
  
  if (headers.indexOf(idHeaderName) === -1) {
    throw new Error('ID header not found')
  }
  
  if (data.length === 0) {data.push(headers)}
  var numberOfColumns = headers.length
  var headerOffsets = getHeaderOffsets()
  
  // {3:{b:[4,5],c:1}}
  
  // [['a','b','c'],[3,4,1],[3,5,1]]
  
  for (var id in objects) {
  
    if (!objects.hasOwnProperty(id)) {continue}
    id = getId(id)    
    var nextObject = objects[id]
    var numberOfRows = getMaximumValueLength(nextObject) // 2

    for (var rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {

      var nextRow = getEmptyRow() // ['','','']
      nextRow[headerOffsets[idHeaderName]] = id // [3,'','']

      for (var header in nextObject) {    
      
        if (!nextObject.hasOwnProperty(header)) {continue}
        var nextValue = nextObject[header]
        nextValue = (isArray(nextValue)) ? (nextValue[rowIndex] || '') : nextValue 
        nextRow[headerOffsets[header]] = nextValue // [3,4,1]
        
      } // for each header in a row
      
      data.push(nextRow)
      
    } // for each row in the longest value array
    
  } // for each ID
  
  return data
  
  // Private Functions
  // -----------------

  /**
   * Some values are arrays, get the length of the longest one
   */

  function getMaximumValueLength(object) {
    var maxLength = 1
    for (var header in object) {
      if (!nextObject.hasOwnProperty(header)) {continue}
      var nextValue = object[header]
      if (isArray(nextValue)) {
        maxLength = (nextValue.length > maxLength) ? nextValue.length : maxLength
      }
    }
    return maxLength
  }

  function getEmptyRow() {
    var row = []
    for (var i = 0; i < numberOfColumns; i++) {row[i] = ''}
    return row
  }

  function getHeaders() {

    if (data.length > 0) {return data[0]}
    var headers = [idHeaderName]
    var exampleObject = objects[Object.keys(objects)[0]]

    for (var header in exampleObject) {
      if (!exampleObject.hasOwnProperty(header)) {continue}
      headers.push(header)
    }

    return headers

  } // SsObjects.addObjectsToArray.getHeaders()

  function getHeaderOffsets() {
    
    var offsets = {}
    
    headers.forEach(function(header, index) {
      offsets[header.trim()] = index
    })
    
    return offsets
    
  } // SsObjects.addObjectsToArray.getHeaderOffsets()

  function getId(oldId) {
    
    if (typeof oldId === 'number') {
    
      return oldId
      
    } else if (typeof oldId === 'string') {     
    
      // Number only works if the whole string is a number, parseInt would 
      // stop once it found non-numeric chars, i.e. oldId may start with a number
      // but contain alphabetic chars
      newId = Number(oldId) 
      
      if (newId !== newId) {// isNan
        return oldId
      } else {
        return newId
      }
      
    } else {
      throw new Error('ID must be a string or number')
    }
    
  } // SsObjects.addObjectsToArray.getId()

} // SsObjects.addObjectsToArray() 

// TODO - Update style for library

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
//
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.

// Originally taken from the Google sample code: 
//
//   https://script.google.com/d/Mg33xUQ0v-ffAw4kUGPlXVXHAGDwXQ1CH/edit?usp=drive_web

function setRowsData (sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders_(headersRange.getValues()[0]);
  
  var data = [];
  
  for (var i = 0; i < objects.length; ++i) {
  
    var values = [];
    
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
    }
    
    data.push(values);
  }
  
  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(), 
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}
   
// getRowsData iterates row by row in the input range and returns an array of objects.
//
// Each object contains all the data for a given row, indexed by its normalized column name.
//
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
//   - dontNormalizeHeaders: optional Boolean, defaults to normalize 
//
// Returns an Array of objects.

// Originally taken from the Google sample code: 
//
//   https://script.google.com/d/Mg33xUQ0v-ffAw4kUGPlXVXHAGDwXQ1CH/edit?usp=drive_web

function getRowsData(sheet, range, columnHeadersRowIndex, dontNormalizeHeaders) {
  
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  
  if (typeof dontNormalizeHeaders === 'undefined' || !dontNormalizeHeaders) {
    headers = normalizeHeaders_(headers)
  }
  
  return getObjects(range.getValues(), headers);
  
  // Private Functions
  // -----------------
  
  // For every row of data in data, generates an object that contains the data. Names of
  // object fields are defined in keys.
  //
  // Arguments:
  //   - data: JavaScript 2d array
  //   - keys: Array of Strings that define the property names for the objects to create
  
  function getObjects(data, keys) {
    
    var objects = [];
    
    for (var i = 0; i < data.length; ++i) {
    
      var object = {};
      var hasData = false;
      
      for (var j = 0; j < data[i].length; ++j) {
      
        var cellData = data[i][j];
        
        if (isCellEmpty(cellData)) {
          continue;
        }
        
        object[keys[j]] = cellData;
        hasData = true;
      }
      
      if (hasData) {
        objects.push(object);
      }
    }
    
    return objects;
    
  } // getRowsData.getObjects()
  
  // Returns true if the cell where cellData was read from is empty.
  // Arguments:
  //   - cellData: string
  
  function isCellEmpty(cellData) {
    
    return typeof(cellData) == "string" && cellData == "";
    
  } // getRowsData.isCellEmpty()
  
} // getRowsData()
   
// Returns an Array of normalized Strings.
//
// Arguments:
//
//   - headers: Array of Strings to normalize

function normalizeHeaders_(headers) {
  
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
  
  // Private Functions
  // -----------------
  
  // Normalizes a string, by removing all alphanumeric characters and using mixed case
  // to separate words. The output will always start with a lower case letter.
  // This function is designed to produce JavaScript object property names.
  //
  // Arguments:
  //   - header: string to normalize
  //
  // Examples:
  //   "First Name" -> "firstName"
  //   "Market Cap (millions) -> "marketCapMillions
  //   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
  
  function normalizeHeader(header) {
    
    var key = "";
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
      var letter = header[i];
      if (letter == " " && key.length > 0) {
        upperCase = true;
        continue;
      }
      //if (!isAlnum(letter)) {
      //  continue;
      //}
      if (key.length == 0 && isDigit(letter)) {
        continue; // first character must be a letter
      }
      if (upperCase) {
        upperCase = false;
        key += letter.toUpperCase();
      } else {
        key += letter.toLowerCase();
      }
    }
    
    return key;
    
    // Private Functions
    // -----------------
    
    // Returns true if the character char is alphabetical, false otherwise.
    
    function isAlnum(char) {
      
      return char >= 'A' && char <= 'Z' ||
        char >= 'a' && char <= 'z' ||
          isDigit(char);
          
    } // isAlnum()
    
    // Returns true if the character char is a digit, false otherwise.
    
    function isDigit(char) {
      
      return char >= '0' && char <= '9';
      
    } // isDigit()
    
  } // normalizeHeader()
  
} // normalizeHeaders_()

function isISODateString_(dateString) {
  if (typeof dateString !== 'string') {return false}
  // E.g. 2020-01-01T12:43:26.000Z
  var regex = /20\d{2}(-|\/)((0[1-9])|(1[0-2]))(-|\/)((0[1-9])|([1-2][0-9])|(3[0-1]))(T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9]):([0-5][0-9]).([0-9][0-9][0-9])(Z)/
  var result = dateString.match(regex)
  return (result !== null && result.length === 18)
}
