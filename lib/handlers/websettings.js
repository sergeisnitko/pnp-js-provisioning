"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
/**
 * Describes the Features Object Handler
 */
class WebSettings extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("WebSettings");
    }
    /**
     * Provisioning features
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web, settings) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            web.update(settings).then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => reject(e));
        });
    }
}
exports.WebSettings = WebSettings;
