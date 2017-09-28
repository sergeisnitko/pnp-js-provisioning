"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const ReplaceTokens = (str) => {
    return str.replace(/{sitecollection}/g, _spPageContextInfo.siteAbsoluteUrl)
        .replace(/{wpgallery}/g, `${_spPageContextInfo.siteAbsoluteUrl}/_catalogs/wp`)
        .replace(/{themegallery}/g, `${_spPageContextInfo.siteAbsoluteUrl}/_catalogs/theme/15`);
};
exports.ReplaceTokens = ReplaceTokens;
const MakeUrlRelative = (absUrl) => {
    return absUrl.replace(`${document.location.protocol}//${document.location.hostname}`, "");
};
exports.MakeUrlRelative = MakeUrlRelative;
