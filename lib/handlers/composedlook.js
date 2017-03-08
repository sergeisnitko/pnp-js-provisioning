"use strict";
const handlerbase_1 = require("./handlerbase");
/**
 * Describes the Composed Look Object Handler
 */
class ComposedLook extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectComposedLook class
     */
    constructor() {
        super("ComposedLook");
    }
    /**
     * Provisioning Composed Look
     *
     * @param object The Composed Look to provision
     */
    ProvisionObjects(web, composedLook) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            web.applyTheme(composedLook.ColorPaletteUrl, composedLook.FontSchemeUrl, composedLook.BackgroundImageUrl, true).then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
}
exports.ComposedLook = ComposedLook;
