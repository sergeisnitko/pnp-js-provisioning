const ReplaceTokens = (str: string): string => {
    return str.replace(/{sitecollection}/g, _spPageContextInfo.siteAbsoluteUrl)
        .replace(/{wpgallery}/g, `${_spPageContextInfo.siteAbsoluteUrl}/_catalogs/wp`)
        .replace(/{themegallery}/g, `${_spPageContextInfo.siteAbsoluteUrl}/_catalogs/theme/15`);
};

const MakeUrlRelative = (absUrl: string): string => {
    return absUrl.replace(`${document.location.protocol}//${document.location.hostname}`, "");
};

export { ReplaceTokens, MakeUrlRelative };
