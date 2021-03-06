var path  = require('path');
const traverse = require('traverse');
var helper = require('./helper');


var preprocessor = {

  /**
   * Execute preprocessor on main, and embedded document
   * @param  {Object}   template
   * @param  {Function} callback
   */
  execute : function (template, callback) {
    if (template === null || template.files === undefined) {
      return callback(null, template);
    }
    for (var i = -1; i < template.embeddings.length; i++) {
      var _mainOrEmbeddedTemplate = template.filename;
      var _parentFilter = '';
      var _fileType = template.extension;
      if (i > -1) {
        // If the current template is an embedded file
        _mainOrEmbeddedTemplate = _parentFilter = template.embeddings[i];
        _fileType = path.extname(_mainOrEmbeddedTemplate).toLowerCase().slice(1);
      }
      switch (_fileType) {
        case 'xlsx':
          preprocessor.convertSharedStringToInlineString(template, _parentFilter);
          // Add function to adjust table range
          preprocessor.adjustTableHeights(template, _parentFilter);
          break;
        case 'ods':
          preprocessor.convertNumberMarkersIntoNumericFormat(template);
          break;
        case 'odt':
          preprocessor.removeSoftPageBreak(template);
          break;
        default:
          break;
      }
    }
    return callback(null, template);
  },

  /**
   * Remove soft page break in ODT document as we modify them
   *
   * Search title "text:use-soft-page-breaks" in
   * http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html#__RefHeading__1415190_253892949
   *
   * TODO: Should we do it in Word?
   *
   * @param  {Object} template (modified)
   * @return {Object}          template
   */
  removeSoftPageBreak : function (template) {
    for (var i = 0; i < template.files.length; i++) {
      var _file = template.files[i];
      if (/content\.xml$/.test(_file.name) === true) {
        _file.data = _file.data
          .replace(/text:use-soft-page-breaks="true"/g, '')
          .replace(/<text:soft-page-break\/>/g, '')
          .replace(/<text:soft-page-break><\/text:soft-page-break>/g, '');
        return template;
      }
    }
    return template;
  },

  /**
   * [XLSX] Convert shared string to inline string in Excel format in order to be compatible with Carbone algorithm
   * @param  {Object} template     (modified)
   * @param  {String} parentFilter apply the transformation on a specific embedded file
   * @return {Object}              the modified template
   */
  convertSharedStringToInlineString : function (template, parentFilter) {
    var _sharedStrings = [];
    var _filesToConvert = [];
    var _sharedStringIndex = -1;
    // parse all files and find shared strings first
    for (var i = 0; i < template.files.length; i++) {
      var _file = template.files[i];
      if (_file.parent === parentFilter) {
        if (/sharedStrings\.xml$/.test(_file.name) === true) {
          _sharedStringIndex = i;
          _sharedStrings = preprocessor.readSharedString(_file.data);
        }
        else if (/\.xml$/.test(_file.name) === true) {
          if (_file.name.indexOf('sheet') !== -1) {
            _filesToConvert.push(_file);
          }
        }
      }
    }
    preprocessor.removeOneFile(template, _sharedStringIndex, parentFilter);
    // once shared string is found, convert files
    for (var f = 0; f < _filesToConvert.length; f++) {
      var _modifiedFile = _filesToConvert[f];
      _modifiedFile.data = preprocessor.removeRowCounterInWorksheet(
        preprocessor.convertToInlineString(_modifiedFile.data, _sharedStrings)
      );
    }
    return template;
  },

  /**
   * Function for xlsx templates.
   * Will look for a string :tableGuide(tableName) that,
   * given the table name and a property to which it's applied,
   * will adjust the ref's in the corresponding tables table
   * @param {Object} template (modified)
   * @param {String} parentFilter apply the transformation on a specific embedded file
   */
  adjustTableHeights : function (template, parentFilter) {
    const _sheets = [];
    const _tables = [];
    const _tableRefs = {};
    for (let i = 0; i < template.files.length; i += 1) {
      const _file = template.files[i];
      // Process only files according to the parentFilter
      if (_file.parent === parentFilter) {
        if (/\/tables\/.*\.xml/.test(_file.name)) {
          _tables.push(_file);
        }
        if (/sheet.+\.xml$/.test(_file.name)) {
          _sheets.push(_file);
        }
      }
    }
    // Process table refs
    for (let i = 0; i < _sheets.length; i += 1) {
      const _sheet = _sheets[i];
      _sheet.data = _sheet.data.replace(/\{[dc]\.(.+?):tableGuide\('?(.+?)'?\)\}/g, function (m, path, tableName) {
        const dataArray = traverse.get(template.data, path.split('.'));
        if (!(dataArray instanceof Array)) {
          throw new Error(':tableGuide should receive an array');
        }
        _tableRefs[tableName] = {
          name   : tableName,
          path   : path,
          length : dataArray.length
        };
        return '';
      });
    }
    // Process tables
    for (let i = 0; i < _tables.length; i += 1) {
      const _table = _tables[i];
      const name = _table.data.match(/name=['"](.+?)['"]/);
      if (name && _tableRefs[name[1]]) {
        _table.data = _table.data.replace(/(ref=['"][A-Z]{1,3})(\d+)(:[A-Z]{1,3})(\d+)(['"])/g, function(m, beg, startIdx, mid, endIdx, end) {
          const start = +startIdx;
          return `${beg}${startIdx}${mid}${start + _tableRefs[name[1]].length}${end}`;
        });
      }
    }
  },

  /**
   * [XLSX] Remove one file in the template and its relations
   * @param  {Object} template
   * @param  {Integer} indexOfFileToRemove index of the file to remove
   * @param  {String} parentFilter         filter to modify only an embedded document
   * @return {}                            it modifies the template directly
   */
  removeOneFile : function (template, indexOfFileToRemove, parentFilter) {
    if (indexOfFileToRemove < 0 || indexOfFileToRemove >= template.files.length) {
      return;
    }
    var _fileToRemove = template.files[indexOfFileToRemove];
    var _dirname  = path.dirname(_fileToRemove.name);
    var _basename = path.basename(_fileToRemove.name);
    template.files.splice(indexOfFileToRemove, 1);
    for (var i = 0; i < template.files.length; i++) {
      var _file = template.files[i];
      if (_file.parent === parentFilter) {
        // remove relations
        if (_dirname + '/_rels/workbook.xml.rels' === _file.name) {
          var _regExp = new RegExp('<Relationship [^>]*Target="' + helper.regexEscape(_basename) + '"[^>]*/>');
          _file.data = _file.data.replace(_regExp, '');
        }
      }
    }
  },

  /**
   * [XLSX] Parse and generate an array of shared string
   * @param  {String} sharedStringXml shared string content
   * @return {Array}                  array
   */
  readSharedString : function (sharedStringXml) {
    var _sharedStrings = [];
    if (sharedStringXml === null || sharedStringXml === undefined) {
      return _sharedStrings;
    }
    var _tagRegex = new RegExp('<si>(.+?)</si>','g');
    var _tag = _tagRegex.exec(sharedStringXml);
    while (_tag !== null) {
      _sharedStrings.push(_tag[1]);
      _tag = _tagRegex.exec(sharedStringXml);
    }
    return _sharedStrings;
  },

  /**
   * [XLSX] Inject shared string in sheets
   * @param  {String} xml           sheets where to insert shared strings
   * @param  {Array} sharedStrings  shared string
   * @return {String}               updated xml
   */
  convertToInlineString : function (xml, sharedStrings) {
    if (typeof(xml) !== 'string') {
      return xml;
    }
    // find all tags which have attribute t="e" (type = functions)
    var _functionXml = xml.replace(/(<(\w)[^>]*>)(<f>.*<\/f>.*?)(<\/\2>)/g, function (m, openTag, tagName, content, closeTag) {
      // get the index of shared string
      var _tab = /<v>(.+?)<\/v>/.exec(content);
      if (_tab instanceof Array && _tab.length > 0) {
        // Just remove the content, it will be recalculated on opening the file
        return openTag + content.replace(/<v>(.+?)<\/v>/, '') + closeTag;
      }
      // if something goes wrong, do nothing
      return m;
    });
    // find all tags which have attribute t="s" (type = shared string)
    var _inlinedXml = _functionXml.replace(/(<(\w)[^>]*t="s"[^>]*>)(.*?)(<\/\2>)/g, function (m, openTag, tagName, content, closeTag) {
      var _newXml = '';
      // get the index of shared string
      var _tab = /<v>(\d+?)<\/v>/.exec(content);
      if (_tab instanceof Array && _tab.length > 0) {
        // replace the index by the string
        var _sharedStringIndex = parseInt(_tab[1], 10);
        var _selectedSharedString = sharedStrings[_sharedStringIndex];
        const _contentAsNumber = /{[d|c][.].*:formatN\(.*\)}/.exec(_selectedSharedString);
        const _contentAsDate = /{[d|c][.].*:formatDExcel\(.*\)}/.exec(_selectedSharedString);
        if (_contentAsNumber instanceof Array && _tab.length > 0) {
          // Convert the marker into number cell type when it using the ':formatN' formatter
          _selectedSharedString = _contentAsNumber[0].replace(/:formatN\(.*\)/, '');
          _newXml = openTag.replace('t="s"', 't="n"');
          _newXml += '<v>' + _selectedSharedString + '</v>';
        }
        else if (_contentAsDate instanceof Array && _tab.length > 0) {
          // Convert the merker into date cell by removing the typ
          _selectedSharedString = _contentAsDate[0].replace(/<\/?t>/, '');
          _newXml = openTag.replace('t="s"', '');
          _newXml += '<v>' + _selectedSharedString + '</v>';
        }
        else {
          // change type of tag to "inline string"
          _newXml = openTag.replace('t="s"', 't="inlineStr"');
          _newXml += '<is>' + _selectedSharedString + '</is>';
        }
        _newXml += closeTag;
        return _newXml;
      }
      // if something goes wrong, do nothing
      return m;
    });
    return _inlinedXml;
  },

  /**
   * [XLSX] Remove row and column counter (r=1, c=A1) in sheet (should be added in post-processing)
   * Carbone Engine cannot update these counter itself
   * @param  {String} xml sheet
   * @return {String}     sheet updated
   */
  removeRowCounterInWorksheet : function (xml) {
    if (typeof(xml) !== 'string') {
      return xml;
    }
    return xml.replace(/<(?:c|row)[^>]*\s(r="\S+")[^>]*>/g, function (m, rowValue) {
      return m.replace(rowValue, '');
    }).replace(/<(?:c|row)[^>]*(spans="\S+")[^>]*>/g, function (m, rowValue) {
      return m.replace(rowValue, '');
    });
  },

  /**
   * @description [ODS] convert number markers with the `formatN()` formatter into numeric format
   *
   * @param {Object} template
   * @return
   */
  convertNumberMarkersIntoNumericFormat : function (template) {
    const _contentFileId = template.files.findIndex(x => x.name === 'content.xml');
    if (_contentFileId > -1 && !!template.files[_contentFileId] === true) {
      template.files[_contentFileId].data = template.files[_contentFileId].data.replace(/<table:table-cell[^<]*>\s*<text:p>[^<]*formatN[^<]*<\/text:p>\s*<\/table:table-cell>/g, function (xml) {
        const _markers = xml.match(/(\{[^{]+?\})/g);
        // we cannot convert to number of there are multiple markers in the same cell
        if (_markers.length !== 1) {
          return xml;
        }
        const _marker = _markers[0].replace(/:formatN\(.*\)/, '');
        xml = xml.replace(/:formatN\(.*\)/, '');
        xml = xml.replace(/office:value-type="string"/, `office:value-type="float" office:value="${_marker}"`);
        xml = xml.replace(/calcext:value-type="string"/, 'calcext:value-type="float"');
        return xml;
      });
    }
    return template;
  }
};

module.exports = preprocessor;