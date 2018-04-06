'use strict';

var RTFParser = require('./rtf-parser.js');
var RTFDocument = require('./rtf-document.js');
var RTFInterpreter = require('./rtf-interpreter.js');

module.exports = parse;
parse.string = parseString;
parse.stream = parseStream;

function parseString(string, cb) {
  parse(cb).end(string);
}

function parseStream(stream, cb) {
  stream.pipe(parse(cb));
}

function parse(cb) {
  var errored = false;
  var errorHandler = function errorHandler(err) {
    if (errored) return;
    errored = true;
    parser.unpipe(interpreter);
    interpreter.end();
    cb(err);
  };
  var document = new RTFDocument();
  var parser = new RTFParser();
  parser.once('error', errorHandler);
  var interpreter = new RTFInterpreter(document);
  interpreter.on('error', errorHandler);
  interpreter.on('finish', function () {
    return cb(null, document);
  });
  parser.pipe(interpreter);
  return parser;
}