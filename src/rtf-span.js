"use strict";

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var RTFSpan = function RTFSpan(opts) {
  _classCallCheck(this, RTFSpan);

  if (!opts) opts = {};
  this.value = opts.value;
  this.style = opts.style || {};
};

module.exports = RTFSpan;