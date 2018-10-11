'use strict';

var Plugin = require('markdown-it-regexp');

module.exports = function fontawesome_plugin(md) {
    // FA4 style.
    md.use(Plugin(
        /:fa-([\w\-]+?)(?:\s+(\d+))?:/,
        function (match, utils) {
            return `<img src="${document.location.origin}/assets/fa4/${utils.escape(match[1])}.png" height="${match[2] || 32}" align="bottom"/>`;
        }
    ));
};
