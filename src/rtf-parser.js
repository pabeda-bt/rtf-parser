'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Transform = require('readable-stream').Transform;

var RTFParser = function (_Transform) {
  _inherits(RTFParser, _Transform);

  function RTFParser() {
    _classCallCheck(this, RTFParser);

    var _this = _possibleConstructorReturn(this, (RTFParser.__proto__ || Object.getPrototypeOf(RTFParser)).call(this, { objectMode: true }));

    _this.text = '';
    _this.controlWord = '';
    _this.controlWordParam = '';
    _this.hexChar = '';
    _this.parserState = _this.parseText;
    _this.char = 0;
    _this.row = 1;
    _this.col = 1;
    return _this;
  }

  _createClass(RTFParser, [{
    key: '_transform',
    value: function _transform(buf, encoding, done) {
      var text = buf.toString('ascii');
      for (var ii = 0; ii < text.length; ++ii) {
        ++this.char;
        if (text[ii] === '\n') {
          ++this.row;
          this.col = 1;
        } else {
          ++this.col;
        }
        this.parserState(text[ii]);
      }
      done();
    }
  }, {
    key: '_flush',
    value: function _flush(done) {
      this.emitText();
      done();
    }
  }, {
    key: 'parseText',
    value: function parseText(char) {
      if (char === '\\') {
        this.parserState = this.parseEscapes;
      } else if (char === '{') {
        this.emitStartGroup();
      } else if (char === '}') {
        this.emitEndGroup();
      } else if (char === '\x0A' || char === '\x0D') {
        // cr/lf are noise chars
      } else {
        this.text += char;
      }
    }
  }, {
    key: 'parseEscapes',
    value: function parseEscapes(char) {
      if (char === '\\' || char === '{' || char === '}') {
        this.text += char;
        this.parserState = this.parseText;
      } else {
        this.parserState = this.parseControlSymbol;
        this.parseControlSymbol(char);
      }
    }
  }, {
    key: 'parseControlSymbol',
    value: function parseControlSymbol(char) {
      if (char === '~') {
        this.text += '\xA0'; // nbsp
        this.parserState = this.parseText;
      } else if (char === '-') {
        this.text += '\xAD'; // soft hyphen
      } else if (char === '_') {
        this.text += '\u2011'; // non-breaking hyphen
      } else if (char === '*') {
        this.emitIgnorable();
        this.parserState = this.parseText;
      } else if (char === "'") {
        this.parserState = this.parseHexChar;
      } else if (char === '|') {
        // formula character
        this.emitFormula();
        this.parserState = this.parseText;
      } else if (char === ':') {
        // subentry in an index entry
        this.emitIndexSubEntry();
        this.parserState = this.parseText;
      } else if (char === '\x0a') {
        this.emitEndParagraph();
        this.parserState = this.parseText;
      } else if (char === '\x0d') {
        this.emitEndParagraph();
        this.parserState = this.parseText;
      } else {
        this.parserState = this.parseControlWord;
        this.parseControlWord(char);
      }
    }
  }, {
    key: 'parseHexChar',
    value: function parseHexChar(char) {
      if (/^[A-Fa-f0-9]$/.test(char)) {
        this.hexChar += char;
        if (this.hexChar.length >= 2) {
          this.emitHexChar();
          this.parserState = this.parseText;
        }
      } else {
        this.emitError('Invalid character "' + char + '" in hex literal.');
        this.parserState = this.parseText;
      }
    }
  }, {
    key: 'parseControlWord',
    value: function parseControlWord(char) {
      if (char === ' ') {
        this.emitControlWord();
        this.parserState = this.parseText;
      } else if (/^[-\d]$/.test(char)) {
        this.parserState = this.parseControlWordParam;
        this.controlWordParam += char;
      } else if (/^[A-Za-z]$/.test(char)) {
        this.controlWord += char;
      } else {
        this.emitControlWord();
        this.parserState = this.parseText;
        this.parseText(char);
      }
    }
  }, {
    key: 'parseControlWordParam',
    value: function parseControlWordParam(char) {
      if (/^\d$/.test(char)) {
        this.controlWordParam += char;
      } else if (char === ' ') {
        this.emitControlWord();
        this.parserState = this.parseText;
      } else {
        this.emitControlWord();
        this.parserState = this.parseText;
        this.parseText(char);
      }
    }
  }, {
    key: 'emitText',
    value: function emitText() {
      if (this.text === '') return;
      this.push({ type: 'text', value: this.text, pos: this.char, row: this.row, col: this.col });
      this.text = '';
    }
  }, {
    key: 'emitControlWord',
    value: function emitControlWord() {
      this.emitText();
      if (this.controlWord === '') {
        this.emitError('empty control word');
      } else {
        this.push({
          type: 'control-word',
          value: this.controlWord,
          param: this.controlWordParam !== '' && Number(this.controlWordParam),
          pos: this.char,
          row: this.row,
          col: this.col
        });
      }
      this.controlWord = '';
      this.controlWordParam = '';
    }
  }, {
    key: 'emitStartGroup',
    value: function emitStartGroup() {
      this.emitText();
      this.push({ type: 'group-start', pos: this.char, row: this.row, col: this.col });
    }
  }, {
    key: 'emitEndGroup',
    value: function emitEndGroup() {
      this.emitText();
      this.push({ type: 'group-end', pos: this.char, row: this.row, col: this.col });
    }
  }, {
    key: 'emitIgnorable',
    value: function emitIgnorable() {
      this.emitText();
      this.push({ type: 'ignorable', pos: this.char, row: this.row, col: this.col });
    }
  }, {
    key: 'emitHexChar',
    value: function emitHexChar() {
      this.emitText();
      this.push({ type: 'hexchar', value: this.hexChar, pos: this.char, row: this.row, col: this.col });
      this.hexChar = '';
    }
  }, {
    key: 'emitError',
    value: function emitError(message) {
      this.emitText();
      this.push({ type: 'error', value: message, row: this.row, col: this.col, char: this.char, stack: new Error().stack });
    }
  }, {
    key: 'emitEndParagraph',
    value: function emitEndParagraph() {
      this.emitText();
      this.push({ type: 'end-paragraph', pos: this.char, row: this.row, col: this.col });
    }
  }]);

  return RTFParser;
}(Transform);

module.exports = RTFParser;