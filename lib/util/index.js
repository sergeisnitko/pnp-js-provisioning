"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const ReplaceTokens = (str) => {
    return str
        .replace(/{sitecollection}/, _spPageContextInfo.siteAbsoluteUrl);
};
exports.ReplaceTokens = ReplaceTokens;
