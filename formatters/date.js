var moment = require('moment');


/**
 * Format dates
 *
 * Since 1.2.0, by default, it considers the input format is "ISO 8601"
 *
 * @exampleContext {"lang":"en"}
 * @example ["20160131", "L"]
 * @example ["20160131", "LL"]
 * @example ["20160131", "LLLL"]
 * @example ["20160131", "dddd"]
 *
 * @exampleContext {"lang":"fr"}
 * @example ["2017-05-10T15:57:23.769561+03:00", "LLLL"]
 * @example ["2017-05-10 15:57:23.769561+03:00", "LLLL"]
 * @example ["20160131", "LLLL"]
 * @example ["20160131", "dddd"]
 *
 * @exampleContext {"lang":"fr"}
 * @example ["20160131", "dddd", "YYYYMMDD"]
 * @example [1410715640, "LLLL", "X" ]
 *
 * @param  {String|Number} d   date to format
 * @param  {String} patternOut output format
 * @param  {String} patternIn  [optional] input format, ISO8601 by default
 * @return {String}            return formatted date
 */
function formatD (d, patternOut, patternIn) {
  if (d !== null && typeof d !== 'undefined') {
    moment.locale(this.lang);
    if (patternIn) {
      return moment(d + '', patternIn).format(patternOut);
    }
    return moment(d + '').format(patternOut);
  }
  return d;
}

/**
 * Convert date to excel format
 * Excel date is stored as the number of days
 * since 1900-01-01. There is a bug in excel, that date 1900-02-29 exists
 * in excel, but it was not a leap year.
 * Time is stored as a number between 0 and .99999 where 0 is 00:00 and .99999 is 23:59
 * @param {Date} d - date to format
 */
function formatDExcel (d) {
  if (!d) return '';
  const excelStartDate = new Date(1900, 0, 0);
  const dMoment = moment(d);
  let result = dMoment.diff(excelStartDate, 'days');
  if (d > new Date('1900-02-28')) {
    result += 1;
  }
  // Get the number of seconds passed today
  const seconds = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
  result += (.99999 * seconds) / 86399;
  return result;
}


/**
 * Format dates
 *
 * @deprecated
 *
 * @exampleContext {"lang":"en"}
 * @example ["20160131", "YYYYMMDD", "L"]
 * @example ["20160131", "YYYYMMDD", "LL"]
 * @example ["20160131", "YYYYMMDD", "LLLL"]
 * @example ["20160131", "YYYYMMDD", "dddd"]
 * @example [1410715640, "X", "LLLL"]
 *
 * @exampleContext {"lang":"fr"}
 * @example ["20160131", "YYYYMMDD", "LLLL"]
 * @example ["20160131", "YYYYMMDD", "dddd"]
 *
 * @param  {String|Number} d   date to format
 * @param  {String} patternIn  input format
 * @param  {String} patternOut output format
 * @return {String}            return formatted date
 */
function convDate (d, patternIn, patternOut) {
  if (d !== null && typeof d !== 'undefined') {
    moment.locale(this.lang);
    return moment(d + '', patternIn).format(patternOut);
  }
  return d;
}




module.exports = {
  formatD      : formatD,
  formatDExcel : formatDExcel,
  convDate     : convDate
};
