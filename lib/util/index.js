"use strict";
const ReplaceTokens = (str) => {
    return str
        .replace(/{sitecollection}/g, _spPageContextInfo.siteAbsoluteUrl);
};
exports.ReplaceTokens = ReplaceTokens;
