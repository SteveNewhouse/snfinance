/*
* SNFinance - useful economic/financial functions for Google Spreadsheet
*/

/**
* Parses financial formatted numbers (i.e. suffixed with "B" or "Billion", etc)
*
* @param {string} the string representing the financial value
* @return {number} the result converted to a number
*/
function parseNum(aString) {
  // TODO: check for "B", "M", "Billion", "Million", "b" "m", etc
  var magnitude = aString.charAt(aString.length-1);
  var value = aString.substring(0, aString.length-1);
  
  if(magnitude == "B") {
     value *= 1000000000;
  }
  else if(magnitude == "M") {
    value *= 1000000;
  }
  
  return value;
}

function _parseDate(input) {
  var parts = input.match(/(\d+)/g);
  // new Date(year, month [, date [, hours[, minutes[, seconds[, ms]]]]])
  return new Date(parts[0], parts[1]-1, parts[2]); // months are 0-based
}

/**
* Calculate the correlation between closing prices two stocks
*
* @param {string} Ticker for stock A
* @param {string} Ticker for stock B
* @param {string} Start date for comparison time window
* @param {string} End date for comparison time window
* @return {number} The calculated correlation between closing prices of stock A and stock B 
*/
function Corr(a, b, startDateStr, endDateStr) {
  var startDate = _parseDate(startDateStr);
  var endDate = _parseDate(endDateStr);
  
  a = FinanceApp.getHistoricalStockInfo(a, startDate, endDate, 1).stockInfo;
  b = FinanceApp.getHistoricalStockInfo(b, startDate, endDate, 1).stockInfo;

  var shortestArrayLength = a.length;
 
  var ab = [];
  var a2 = [];
  var b2 = [];
  
  for(var i=0; i<shortestArrayLength; i++)
  {
      ab.push(a[i].close * b[i].close);
      a2.push(a[i].close * a[i].close);
      b2.push(b[i].close * b[i].close);
  }
  
  var sum_a = 0;
  var sum_b = 0;
  var sum_ab = 0;
  var sum_a2 = 0;
  var sum_b2 = 0;
  
  for(var i=0; i < shortestArrayLength; i++)
  {
      sum_a += a[i].close;
      sum_b += b[i].close;
      sum_ab += ab[i];
      sum_a2 += a2[i];
      sum_b2 += b2[i];
  }
  
  var step1 = (shortestArrayLength * sum_ab) - (sum_a * sum_b);
  var step2 = (shortestArrayLength * sum_a2) - (sum_a * sum_a);
  var step3 = (shortestArrayLength * sum_b2) - (sum_b * sum_b);
  var step4 = Math.sqrt(step2 * step3);
  var answer = step1 / step4;
  
  return answer;
}


/**
* Get YahooFinance! EBITDA for a stock
*
* @param {string} Ticker for stock
* @return {number} The EBITA value
*/
function EBITDA(ticker) {
  var result = YahooFinance(ticker, "j4");
  return parseNum(result);
}

/**
* Get YahooFinance! Price/Earnings Growth value for a stock
*
* @param {string} Ticker for stock
* @return {number} The PEG value
*/
function PEG(ticker) {
  var result = YahooFinance(ticker, "r5");
  return result;
}

// http://www.gummy-stuff.org/Yahoo-data.htm

/**
* Convienence function to access YahooFinance! data for a stock
*
* @param {string} Ticker for stock
* @param {string} Field code for the desired data point
* @return {number} The data point requested
*/
function YahooFinance(ticker, fieldCode) {
  var yahooUrl = "http://finance.yahoo.com/d/quotes.csv?s=" + ticker + "&f=" + fieldCode;
  
  _randSleep();
  var response = UrlFetchApp.fetch(yahooUrl);
  if(response.getResponseCode() != 200) {
    return("Error: data fetch failed. (Error code:" + response.getResponseCode() + ")");  // return the error message
  }

  var result = response.getContentText();
  result = result.trim();
  return result;
}

/**
* Get YahooFinance! 10 year risk-free-rate
*
* @return {number} 10 year risk-free-rate
*/
function RFR10() {
  var result = YahooFinance("%5ETNX", "l1");
  return result/100;
}

/**
* Get WorldBank GDP number for a country by year 
*
* @param {string} Country code
* @param {number} Year
* @return {number} WorldBank GDP value
*/
function GDP(country, year) {
  var url = "http://api.worldbank.org/en/countries/" + country + "/indicators/NY.GDP.MKTP.CD?date=" + year + "&format=json";
  
  _randSleep();
  var response = UrlFetchApp.fetch(url);
  if(response.getResponseCode() != 200) {
    return("Error: data fetch failed. (Error code:" + response.getResponseCode() + ")");  // return the error message
  }

  var response = response.getContentText();
  response = response.trim();
  
  var jsonResponse = Utilities.jsonParse(response);

  for(var i = 0; i < jsonResponse[1].length ; i++) {
    if(jsonResponse[1][i].date == year) {
      return parseInt(jsonResponse[1][i].value);
    }
  }
  
  return "Not Found";
}

/**
* Get WorldBank Unemployment number for a country by year 
*
* @param {string} Country code
* @param {number} Year
* @return {number} WorldBank Unemployment number
*/
function Unemployment(country, year) {
  var url = "http://api.worldbank.org/en/countries/" + country + "/indicators/SL.UEM.TOTL.ZS?date=" + year + "&format=json";
  
  _randSleep();
  var response = UrlFetchApp.fetch(url);
  if(response.getResponseCode() != 200) {
    return("Error: data fetch failed. (Error code:" + response.getResponseCode() + ")");  // return the error message
  }

  var response = response.getContentText();
  response = response.trim();
  
  var jsonResponse = Utilities.jsonParse(response);

  if(jsonResponse[1][0].date == year) {
    return parseFloat(jsonResponse[1][0].value);
  }  
  return "Not Found";
}

/**
* Convenience function for accessing WorldBank indicator data by country & year
*
* @param {string} Indicator code
* @param {string} Country code
* @param {number} Year
* @return {number} WorldBank indicator data value
*/
function Worldbank(indicator, country, year) {
  var url = "http://api.worldbank.org/en/countries/" + country + "/indicators/" + indicator + "?date=" + year + "&format=json";
  
  _randSleep();
  var response = UrlFetchApp.fetch(url);
  if(response.getResponseCode() != 200) {
    return("Error: data fetch failed. (Error code:" + response.getResponseCode() + ")");  // return the error message
  }

  var response = response.getContentText();
  response = response.trim();
  
  var jsonResponse = Utilities.jsonParse(response);

  if(jsonResponse[1][0].date == year) {
    return parseFloat(jsonResponse[1][0].value);
  }  
  return "Not Found";
}

function _randSleep() {
  var randnumber = Math.random()*10000 + 1000;
  Utilities.sleep(randnumber);
  Utilities.sleep(randnumber);
}