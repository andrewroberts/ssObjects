// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO

// SsObjects.gs
// ============

// Test Sheet: https://docs.google.com/spreadsheets/d/1tAYu2STTAebhwq67LArlsVH5k1QmIMPmOxuqI7HYWrY/edit#gid=0

/**
 * Convert table of data, with the headers in the first row
 * into an object, or add to an existing object:
 *
 *
 * [
 *   [id_header,     header2,     header3, ...],
 *   [row1_id_value, row1_value2, row1_value3, ...],
 *   [row2_id_value, row2_value2, row2_value3, ...],
 * }
 * 
 * =>
 *
 * {
 *   row1_id_value: {
 *     [header2]: row1_value2,
 *     [header3]: row1_value3
 *   }, 
 *   row2_id_value: {
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
    
  var headers = data.shift()
  
  data.forEach(function(row) {
    
    var id = null
    
    row.forEach(function(nextValue, index) {
      
      var nextHeader = headers[index].trim()
      
      if (nextHeader === idHeaderName) {
        
        id = nextValue
        
        if (objects[id] === undefined) {          
          objects[id] = {}                
        }
        
      } else {
        
        nextValue = nextValue || ''
        
        if (nextValue instanceof Date || isISODateString_(nextValue)) {
          nextValue = new Date(nextValue)
        }
        
        if (!id) {
          throw new Error('No ID "' + idHeaderName + '" found for this object yet')
        }
        
        if (objects[id][nextHeader] !== undefined) {
          
          var existingValue = objects[id][nextHeader]
          
          if (existingValue != nextValue) {
            
            throw new Error(
              'Updating [%s][%s] from "%s" to "%s"', 
              id, 
              nextHeader, 
              existingValue.toString(), 
              nextValue)
          }
        }
        
        objects[id][nextHeader] = nextValue
      }       
      
    }) // for each cell
    
  }) // for each row
  
  return objects     
}

/**
 * Convert an object into a 2D table of data, with the headers in the first row,
 * or add to an existing table:
 *
 * {
 *   row1_id_value: {
 *     [header2]: row1_value2,
 *     [header3]: row1_value3
 *   }, 
 *   row2_id_value: {
 *     [header2]: row2_value2,
 *     [header3]: row2_value3
 *   },    
 * }
 *
 * => 
 *
 * [
 *   [id_header,     header2,     header3, ...],
 *   [row1_id_value, row1_value2, row1_value3, ...],
 *   [row2_id_value, row2_value2, row2_value3, ...],
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
  
  for (var id in objects) {
  
    if (!objects.hasOwnProperty(id)) {continue}    
    var nextRow = getEmptyRow()
    nextRow[0] = id
    var nextobject = objects[id]
    
    for (var header in nextobject) {
      if (!nextobject.hasOwnProperty(header)) {continue}
      nextRow[headerOffsets[header]] = nextobject[header]
    }
    
    data.push(nextRow)
  }
  
  return data
  
  // Private Functions
  // -----------------

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

function setRowsData (sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders_(headersRange.getValues()[0]);
  
  var data = [];
  
  for (var i = 0; i < objects.length; ++i) {
    var values = []
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

