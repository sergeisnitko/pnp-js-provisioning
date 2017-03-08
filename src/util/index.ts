const ReplaceTokens = (str: string): string => {
    return str
        .replace(/{sitecollection}/, _spPageContextInfo.siteAbsoluteUrl);
};

export { ReplaceTokens };
