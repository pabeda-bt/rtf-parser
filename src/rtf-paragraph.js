"use strict";

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var RTFParagraph = function RTFParagraph(opts) {
  _classCallCheck(this, RTFParagraph);

  if (!opts) opts = {};
  this.style = opts.style || {};
  this.content = [];
};

module.exports = RTFParagraph;