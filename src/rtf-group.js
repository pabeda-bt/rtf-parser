"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var RTFGroup = function () {
  function RTFGroup(parent) {
    _classCallCheck(this, RTFGroup);

    this.parent = parent;
    this.content = [];
    this.fonts = [];
    this.colors = [];
    this.style = {};
    this.ignorable = null;
  }

  _createClass(RTFGroup, [{
    key: "get",
    value: function get(name) {
      return this[name] != null ? this[name] : this.parent.get(name);
    }
  }, {
    key: "getFont",
    value: function getFont(num) {
      return this.fonts[num] != null ? this.fonts[num] : this.parent.getFont(num);
    }
  }, {
    key: "getColor",
    value: function getColor(num) {
      return this.colors[num] != null ? this.colors[num] : this.parent.getFont(num);
    }
  }, {
    key: "getStyle",
    value: function getStyle(name) {
      if (!name) return Object.assign({}, this.parent.getStyle(), this.style);
      return this.style[name] != null ? this.style[name] : this.parent.getStyle(name);
    }
  }, {
    key: "resetStyle",
    value: function resetStyle() {
      this.style = {};
    }
  }, {
    key: "addContent",
    value: function addContent(node) {
      node.style = Object.assign({}, this.getStyle());
      node.style.font = this.getFont(node.style.font);
      node.style.foreground = this.getColor(node.style.foreground);
      node.style.background = this.getColor(node.style.background);
      this.content.push(node);
    }
  }]);

  return RTFGroup;
}();

module.exports = RTFGroup;