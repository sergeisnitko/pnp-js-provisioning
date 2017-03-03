"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Features Object Handler
 */
class Lists extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Lists");
    }
    /**
     * Provisioning lists
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web, lists) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve()).then(() => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
    processList(web, list) {
        return web.lists.ensure(list.Title, list.Description, list.Template, list.ContentTypesEnabled, list.AdditionalSettings).then(result => {
            if (result.created) {
                sp_pnp_js_1.Logger.log({ data: result.list, level: sp_pnp_js_1.LogLevel.Info, message: `List ${list.Title} created successfully.` });
            }
            // here we would do things like add fields, apply content types, etc
        });
    }
}
exports.Lists = Lists;
