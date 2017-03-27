const ReplaceTokens = (str: string): string => {
    return str.replace(/{sitecollection}/g, _spPageContextInfo.siteAbsoluteUrl);
};

const MakeUrlRelative = (absUrl: string): string => {
    return absUrl.replace(`${document.location.protocol}//${document.location.hostname}`, "");
}

export { ReplaceTokens, MakeUrlRelative };
