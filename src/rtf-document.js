'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _get = function get(object, property, receiver) { if (object === null) object = Function.prototype; var desc = Object.getOwnPropertyDescriptor(object, property); if (desc === undefined) { var parent = Object.getPrototypeOf(object); if (parent === null) { return undefined; } else { return get(parent, property, receiver); } } else if ("value" in desc) { return desc.value; } else { var getter = desc.get; if (getter === undefined) { return undefined; } return getter.call(receiver); } };

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var RTFGroup = require('./rtf-group.js');
var RTFParagraph = require('./rtf-paragraph.js');

var RTFDocument = function (_RTFGroup) {
  _inherits(RTFDocument, _RTFGroup);

  function RTFDocument() {
    _classCallCheck(this, RTFDocument);

    var _this = _possibleConstructorReturn(this, (RTFDocument.__proto__ || Object.getPrototypeOf(RTFDocument)).call(this));

    _this.charset = 'ASCII';
    _this.ignorable = false;
    _this.marginLeft = 1800;
    _this.marginRight = 1800;
    _this.marginBottom = 1440;
    _this.marginTop = 1440;
    _this.style = {
      font: 0,
      fontSize: 24,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      foreground: null,
      background: null,
      firstLineIndent: 0,
      indent: 0,
      align: 'left',
      valign: 'normal'
    };
    return _this;
  }

  _createClass(RTFDocument, [{
    key: 'get',
    value: function get(name) {
      return this[name];
    }
  }, {
    key: 'getFont',
    value: function getFont(num) {
      return this.fonts[num];
    }
  }, {
    key: 'getColor',
    value: function getColor(num) {
      return this.colors[num];
    }
  }, {
    key: 'getStyle',
    value: function getStyle(name) {
      if (!name) return this.style;
      return this.style[name];
    }
  }, {
    key: 'addContent',
    value: function addContent(node) {
      if (node instanceof RTFParagraph) {
        while (this.content.length && !(this.content[this.content.length - 1] instanceof RTFParagraph)) {
          node.content.unshift(this.content.pop());
        }
        _get(RTFDocument.prototype.__proto__ || Object.getPrototypeOf(RTFDocument.prototype), 'addContent', this).call(this, node);
        if (node.content.length) {
          var initialStyle = node.content[0].style;
          var style = {};
          style.font = this.getFont(initialStyle.font);
          style.foreground = this.getColor(initialStyle.foreground);
          style.background = this.getColor(initialStyle.background);
          var _iteratorNormalCompletion = true;
          var _didIteratorError = false;
          var _iteratorError = undefined;

          try {
            for (var _iterator = Object.keys(initialStyle)[Symbol.iterator](), _step; !(_iteratorNormalCompletion = (_step = _iterator.next()).done); _iteratorNormalCompletion = true) {
              var prop = _step.value;

              if (initialStyle[prop] == null) continue;
              var match = true;
              var _iteratorNormalCompletion2 = true;
              var _didIteratorError2 = false;
              var _iteratorError2 = undefined;

              try {
                for (var _iterator2 = node.content[Symbol.iterator](), _step2; !(_iteratorNormalCompletion2 = (_step2 = _iterator2.next()).done); _iteratorNormalCompletion2 = true) {
                  var span = _step2.value;

                  if (initialStyle[prop] !== span.style[prop]) {
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

              if (match) style[prop] = initialStyle[prop];
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

          node.style = style;
        }
      } else {
        _get(RTFDocument.prototype.__proto__ || Object.getPrototypeOf(RTFDocument.prototype), 'addContent', this).call(this, node);
      }
    }
  }]);

  return RTFDocument;
}(RTFGroup);

module.exports = RTFDocument;