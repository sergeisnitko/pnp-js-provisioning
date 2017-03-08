"use strict";
const ReplaceTokens = (str) => {
    return str
        .replace(/{sitecollection}/, _spPageContextInfo.siteAbsoluteUrl);
};
exports.ReplaceTokens = ReplaceTokens;
