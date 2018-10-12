/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

'use strict';

const Plugin = require('markdown-it-regexp');

module.exports = function(md) {
    md.use(Plugin(
        /:fa-([\w\-]+?)(?:\s+(\d+))?:/,
        function (match, utils) {
            return `<img src="${document.location.origin}/assets/fa4/${utils.escape(match[1])}.png" height="${match[2] || 32}" align="bottom"/>`;
        }
    ));
};
