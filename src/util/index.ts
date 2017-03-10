const ReplaceTokens = (str: string): string => {
    return str
        .replace(/{sitecollection}/g, _spPageContextInfo.siteAbsoluteUrl);
};

export { ReplaceTokens };
