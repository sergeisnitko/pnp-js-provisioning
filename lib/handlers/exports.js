"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const composedlook_1 = require("./composedlook");
const customactions_1 = require("./customactions");
const features_1 = require("./features");
const websettings_1 = require("./websettings");
const navigation_1 = require("./navigation");
const lists_1 = require("./lists");
const files_1 = require("./files");
exports.DefaultHandlerMap = {
    ComposedLook: new composedlook_1.ComposedLook(),
    CustomActions: new customactions_1.CustomActions(),
    Features: new features_1.Features(),
    Lists: new lists_1.Lists(),
    Navigation: new navigation_1.Navigation(),
    WebSettings: new websettings_1.WebSettings(),
    Files: new files_1.Files(),
};
exports.DefaultHandlerSort = {
    WebSettings: 1,
    Features: 2,
    Files: 7,
    Lists: 3,
    Navigation: 4,
    CustomActions: 5,
    ComposedLook: 6,
};
