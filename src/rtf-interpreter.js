'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var assert = require('assert');
var util = require('util');

var Writable = require('readable-stream').Writable;
var RTFGroup = require('./rtf-group.js');
var RTFParagraph = require('./rtf-paragraph.js');
var RTFSpan = require('./rtf-span.js');
var iconv = require('iconv-lite');

var availableCP = [437, 737, 775, 850, 852, 853, 855, 857, 858, 860, 861, 863, 865, 866, 869, 1125, 1250, 1251, 1252, 1253, 1254, 1257];
var codeToCP = {
  0: 'ASCII',
  77: 'MacRoman',
  128: 'SHIFT_JIS',
  129: 'CP949', // Hangul
  130: 'JOHAB',
  134: 'CP936', // GB2312 simplified chinese
  136: 'BIG5',
  161: 'CP1253', // greek
  162: 'CP1254', // turkish
  163: 'CP1258', // vietnamese
  177: 'CP862', // hebrew
  178: 'CP1256', // arabic
  186: 'CP1257', // baltic
  204: 'CP1251', // russian
  222: 'CP874', // thai
  238: 'CP238', // eastern european
  254: 'CP437' // PC-437
};

var RTFInterpreter = function (_Writable) {
  _inherits(RTFInterpreter, _Writable);

  function RTFInterpreter(document) {
    _classCallCheck(this, RTFInterpreter);

    var _this = _possibleConstructorReturn(this, (RTFInterpreter.__proto__ || Object.getPrototypeOf(RTFInterpreter)).call(this, { objectMode: true }));

    _this.doc = document;
    _this.parserState = _this.parseTop;
    _this.groupStack = [];
    _this.group = null;
    _this.once('prefinish', function () {
      return _this.finisher();
    });
    return _this;
  }

  _createClass(RTFInterpreter, [{
    key: '_write',
    value: function _write(cmd, encoding, done) {
      var method = 'cmd$' + cmd.type.replace(/-(.)/g, function (_, char) {
        return char.toUpperCase();
      });
      if (this[method]) {
        this[method](cmd);
      } else {
        process.emit('error', 'Unknown RTF command ' + cmd.type + ', tried ' + method);
      }
      done();
    }
  }, {
    key: 'finisher',
    value: function finisher() {
      while (this.groupStack.length) {
        this.cmd$groupEnd();
      }var initialStyle = this.doc.content[0].style;
      var _iteratorNormalCompletion = true;
      var _didIteratorError = false;
      var _iteratorError = undefined;

      try {
        for (var _iterator = Object.keys(this.doc.style)[Symbol.iterator](), _step; !(_iteratorNormalCompletion = (_step = _iterator.next()).done); _iteratorNormalCompletion = true) {
          var prop = _step.value;

          var match = true;
          var _iteratorNormalCompletion2 = true;
          var _didIteratorError2 = false;
          var _iteratorError2 = undefined;

          try {
            for (var _iterator2 = this.doc.content[Symbol.iterator](), _step2; !(_iteratorNormalCompletion2 = (_step2 = _iterator2.next()).done); _iteratorNormalCompletion2 = true) {
              var para = _step2.value;

              if (initialStyle[prop] !== para.style[prop]) {
                match = false;
                break;
              }
            }
          } catch (err) {
            _didIteratorError2 = true;
            _iteratorError2 = err;
          } finally {
            try {
              if (!_iteratorNormalCompletion2 && _iterator2.return) {
                _iterator2.return();
              }
            } finally {
              if (_didIteratorError2) {
                throw _iteratorError2;
              }
            }
          }

          if (match) this.doc.style[prop] = initialStyle[prop];
        }
      } catch (err) {
        _didIteratorError = true;
        _iteratorError = err;
      } finally {
        try {
          if (!_iteratorNormalCompletion && _iterator.return) {
            _iterator.return();
          }
        } finally {
          if (_didIteratorError) {
            throw _iteratorError;
          }
        }
      }
    }
  }, {
    key: 'cmd$groupStart',
    value: function cmd$groupStart() {
      if (this.group) this.groupStack.push(this.group);
      this.group = new RTFGroup(this.group || this.doc);
    }
  }, {
    key: 'cmd$ignorable',
    value: function cmd$ignorable() {
      this.group.ignorable = true;
    }
  }, {
    key: 'cmd$endParagraph',
    value: function cmd$endParagraph() {
      this.group.addContent(new RTFParagraph());
    }
  }, {
    key: 'cmd$groupEnd',
    value: function cmd$groupEnd() {
      var endingGroup = this.group;
      this.group = this.groupStack.pop();
      var doc = this.group || this.doc;
      if (endingGroup instanceof FontTable) {
        doc.fonts = endingGroup.table;
      } else if (endingGroup instanceof ColorTable) {
        doc.colors = endingGroup.table;
      } else if (endingGroup !== this.doc && !endingGroup.get('ignorable')) {
        var _iteratorNormalCompletion3 = true;
        var _didIteratorError3 = false;
        var _iteratorError3 = undefined;

        try {
          for (var _iterator3 = endingGroup.content[Symbol.iterator](), _step3; !(_iteratorNormalCompletion3 = (_step3 = _iterator3.next()).done); _iteratorNormalCompletion3 = true) {
            var item = _step3.value;

            doc.addContent(item);
          }
        } catch (err) {
          _didIteratorError3 = true;
          _iteratorError3 = err;
        } finally {
          try {
            if (!_iteratorNormalCompletion3 && _iterator3.return) {
              _iterator3.return();
            }
          } finally {
            if (_didIteratorError3) {
              throw _iteratorError3;
            }
          }
        }

        process.emit('debug', 'GROUP END', endingGroup.type, endingGroup.get('ignorable'));
      }
    }
  }, {
    key: 'cmd$text',
    value: function cmd$text(cmd) {
      this.group.addContent(new RTFSpan(cmd));
    }
  }, {
    key: 'cmd$controlWord',
    value: function cmd$controlWord(cmd) {
      if (!this.group.type) this.group.type = cmd.value;
      var method = 'ctrl$' + cmd.value.replace(/-(.)/g, function (_, char) {
        return char.toUpperCase();
      });
      if (this[method]) {
        this[method](cmd.param);
      } else {
        if (!this.group.get('ignorable')) process.emit('debug', method, cmd.param);
      }
    }
  }, {
    key: 'cmd$hexchar',
    value: function cmd$hexchar(cmd) {
      this.group.addContent(new RTFSpan({
        value: iconv.decode(Buffer.from(cmd.value, 'hex'), this.group.get('charset'))
      }));
    }
  }, {
    key: 'ctrl$rtf',
    value: function ctrl$rtf() {
      this.group = this.doc;
    }

    // alignment

  }, {
    key: 'ctrl$qc',
    value: function ctrl$qc() {
      this.group.style.align = 'center';
    }
  }, {
    key: 'ctrl$qj',
    value: function ctrl$qj() {
      this.group.style.align = 'justify';
    }
  }, {
    key: 'ctrl$ql',
    value: function ctrl$ql() {
      this.group.style.align = 'left';
    }
  }, {
    key: 'ctrl$qr',
    value: function ctrl$qr() {
      this.group.style.align = 'right';
    }

    // text direction

  }, {
    key: 'ctrl$rtlch',
    value: function ctrl$rtlch() {
      this.group.style.dir = 'rtl';
    }
  }, {
    key: 'ctrl$ltrch',
    value: function ctrl$ltrch() {
      this.group.style.dir = 'ltr';
    }

    // general style

  }, {
    key: 'ctrl$par',
    value: function ctrl$par() {
      this.group.addContent(new RTFParagraph());
    }
  }, {
    key: 'ctrl$pard',
    value: function ctrl$pard() {
      this.group.resetStyle();
    }
  }, {
    key: 'ctrl$plain',
    value: function ctrl$plain() {
      this.group.style.fontSize = this.doc.getStyle('fontSize');
      this.group.style.bold = this.doc.getStyle('bold');
      this.group.style.italic = this.doc.getStyle('italic');
      this.group.style.underline = this.doc.getStyle('underline');
    }
  }, {
    key: 'ctrl$b',
    value: function ctrl$b(set) {
      this.group.style.bold = set !== 0;
    }
  }, {
    key: 'ctrl$i',
    value: function ctrl$i(set) {
      this.group.style.italic = set !== 0;
    }
  }, {
    key: 'ctrl$u',
    value: function ctrl$u(num) {
      var charBuf = Buffer.alloc ? Buffer.alloc(2) : new Buffer(2);
      charBuf.writeUInt16LE(num, 0);
      this.group.addContent(new RTFSpan({ value: iconv.decode(charBuf, 'ucs2') }));
    }
  }, {
    key: 'ctrl$super',
    value: function ctrl$super() {
      this.group.style.valign = 'super';
    }
  }, {
    key: 'ctrl$sub',
    value: function ctrl$sub() {
      this.group.style.valign = 'sub';
    }
  }, {
    key: 'ctrl$nosupersub',
    value: function ctrl$nosupersub() {
      this.group.style.valign = 'normal';
    }
  }, {
    key: 'ctrl$strike',
    value: function ctrl$strike(set) {
      this.group.style.strikethrough = set !== 0;
    }
  }, {
    key: 'ctrl$ul',
    value: function ctrl$ul(set) {
      this.group.style.underline = set !== 0;
    }
  }, {
    key: 'ctrl$ulnone',
    value: function ctrl$ulnone(set) {
      this.group.style.underline = false;
    }
  }, {
    key: 'ctrl$fi',
    value: function ctrl$fi(value) {
      this.group.style.firstLineIndent = value;
    }
  }, {
    key: 'ctrl$cufi',
    value: function ctrl$cufi(value) {
      this.group.style.firstLineIndent = value * 100;
    }
  }, {
    key: 'ctrl$li',
    value: function ctrl$li(value) {
      this.group.style.indent = value;
    }
  }, {
    key: 'ctrl$lin',
    value: function ctrl$lin(value) {
      this.group.style.indent = value;
    }
  }, {
    key: 'ctrl$culi',
    value: function ctrl$culi(value) {
      this.group.style.indent = value * 100;
    }

    // encodings

  }, {
    key: 'ctrl$ansi',
    value: function ctrl$ansi() {
      this.group.charset = 'ASCII';
    }
  }, {
    key: 'ctrl$mac',
    value: function ctrl$mac() {
      this.group.charset = 'MacRoman';
    }
  }, {
    key: 'ctrl$pc',
    value: function ctrl$pc() {
      this.group.charset = 'CP437';
    }
  }, {
    key: 'ctrl$pca',
    value: function ctrl$pca() {
      this.group.charset = 'CP850';
    }
  }, {
    key: 'ctrl$ansicpg',
    value: function ctrl$ansicpg(codepage) {
      if (availableCP.indexOf(codepage) === -1) {
        this.emit('error', new Error('Codepage ' + codepage + ' is not available.'));
      } else {
        this.group.charset = 'CP' + codepage;
      }
    }

    // fonts

  }, {
    key: 'ctrl$fonttbl',
    value: function ctrl$fonttbl() {
      this.group = new FontTable(this.group.parent);
    }
  }, {
    key: 'ctrl$f',
    value: function ctrl$f(num) {
      if (this.group instanceof FontTable) {
        this.group.currentFont = this.group.table[num] = new Font();
      } else {
        this.group.style.font = num;
      }
    }
  }, {
    key: 'ctrl$fnil',
    value: function ctrl$fnil() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'nil';
      }
    }
  }, {
    key: 'ctrl$froman',
    value: function ctrl$froman() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'roman';
      }
    }
  }, {
    key: 'ctrl$fswiss',
    value: function ctrl$fswiss() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'swiss';
      }
    }
  }, {
    key: 'ctrl$fmodern',
    value: function ctrl$fmodern() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'modern';
      }
    }
  }, {
    key: 'ctrl$fscript',
    value: function ctrl$fscript() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'script';
      }
    }
  }, {
    key: 'ctrl$fdecor',
    value: function ctrl$fdecor() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'decor';
      }
    }
  }, {
    key: 'ctrl$ftech',
    value: function ctrl$ftech() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'tech';
      }
    }
  }, {
    key: 'ctrl$fbidi',
    value: function ctrl$fbidi() {
      if (this.group instanceof FontTable) {
        this.group.currentFont.family = 'bidi';
      }
    }
  }, {
    key: 'ctrl$fcharset',
    value: function ctrl$fcharset(code) {
      if (this.group instanceof FontTable) {
        var charset = null;
        if (code === 1) {
          charset = this.group.get('charset');
        } else {
          charset = codeToCP[code];
        }
        if (charset == null) {
          return this.emit('error', new Error('Unsupported charset code #' + code));
        }
        this.group.currentFont.charset = charset;
      }
    }
  }, {
    key: 'ctrl$fprq',
    value: function ctrl$fprq(pitch) {
      if (this.group instanceof FontTable) {
        this.group.currentFont.pitch = pitch;
      }
    }

    // colors

  }, {
    key: 'ctrl$colortbl',
    value: function ctrl$colortbl() {
      this.group = new ColorTable(this.group.parent);
    }
  }, {
    key: 'ctrl$red',
    value: function ctrl$red(value) {
      if (this.group instanceof ColorTable) {
        this.group.red = value;
      }
    }
  }, {
    key: 'ctrl$blue',
    value: function ctrl$blue(value) {
      if (this.group instanceof ColorTable) {
        this.group.blue = value;
      }
    }
  }, {
    key: 'ctrl$green',
    value: function ctrl$green(value) {
      if (this.group instanceof ColorTable) {
        this.group.green = value;
      }
    }
  }, {
    key: 'ctrl$cf',
    value: function ctrl$cf(value) {
      this.group.style.foreground = value;
    }
  }, {
    key: 'ctrl$cb',
    value: function ctrl$cb(value) {
      this.group.style.background = value;
    }
  }, {
    key: 'ctrl$fs',
    value: function ctrl$fs(value) {
      this.group.style.fontSize = value;
    }

    // margins

  }, {
    key: 'ctrl$margl',
    value: function ctrl$margl(value) {
      this.doc.marginLeft = value;
    }
  }, {
    key: 'ctrl$margr',
    value: function ctrl$margr(value) {
      this.doc.marginRight = value;
    }
  }, {
    key: 'ctrl$margt',
    value: function ctrl$margt(value) {
      this.doc.marginTop = value;
    }
  }, {
    key: 'ctrl$margb',
    value: function ctrl$margb(value) {
      this.doc.marginBottom = value;
    }

    // unsupported (and we need to ignore content)

  }, {
    key: 'ctrl$stylesheet',
    value: function ctrl$stylesheet(value) {
      this.group.ignorable = true;
    }
  }, {
    key: 'ctrl$info',
    value: function ctrl$info(value) {
      this.group.ignorable = true;
    }
  }, {
    key: 'ctrl$mmathPr',
    value: function ctrl$mmathPr(value) {
      this.group.ignorable = true;
    }
  }]);

  return RTFInterpreter;
}(Writable);

var FontTable = function (_RTFGroup) {
  _inherits(FontTable, _RTFGroup);

  function FontTable(parent) {
    _classCallCheck(this, FontTable);

    var _this2 = _possibleConstructorReturn(this, (FontTable.__proto__ || Object.getPrototypeOf(FontTable)).call(this, parent));

    _this2.table = [];
    _this2.currentFont = { family: 'roman', charset: 'ASCII', name: 'Serif' };
    return _this2;
  }

  _createClass(FontTable, [{
    key: 'addContent',
    value: function addContent(text) {
      this.currentFont.name = text.value.replace(/;\s*$/, '');
    }
  }]);

  return FontTable;
}(RTFGroup);

var Font = function Font() {
  _classCallCheck(this, Font);

  this.family = null;
  this.charset = null;
  this.name = null;
  this.pitch = 0;
};

var ColorTable = function (_RTFGroup2) {
  _inherits(ColorTable, _RTFGroup2);

  function ColorTable(parent) {
    _classCallCheck(this, ColorTable);

    var _this3 = _possibleConstructorReturn(this, (ColorTable.__proto__ || Object.getPrototypeOf(ColorTable)).call(this, parent));

    _this3.table = [];
    _this3.red = 0;
    _this3.blue = 0;
    _this3.green = 0;
    return _this3;
  }

  _createClass(ColorTable, [{
    key: 'addContent',
    value: function addContent(text) {
      assert(text.value === ';', 'got: ' + util.inspect(text));
      this.table.push({
        red: this.red,
        blue: this.blue,
        green: this.green
      });
      this.red = 0;
      this.blue = 0;
      this.green = 0;
    }
  }]);

  return ColorTable;
}(RTFGroup);

module.exports = RTFInterpreter;