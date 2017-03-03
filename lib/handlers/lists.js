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
        return new Promise((resolve, reject) => {
            web.lists.ensure(list.Title, list.Description, list.Template, list.ContentTypesEnabled, list.AdditionalSettings).then(result => {
                if (result.created) {
                    sp_pnp_js_1.Logger.log({ data: result.list, level: sp_pnp_js_1.LogLevel.Info, message: `List ${list.Title} created successfully.` });
                }
                this.processContentTypeBindings(result.list, list.ContentTypeBindings).then(resolve, reject);
            });
        });
    }
    processContentTypeBindings(list, contentTypeBindings) {
        return new Promise((resolve, reject) => {
            contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(list, ct)), Promise.resolve()).then(() => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
    processContentTypeBinding(list, contentTypeBinding) {
        return new Promise((resolve, reject) => {
            list.contentTypes.addAvailableContentType(contentTypeBinding.Id).then(result => {
                sp_pnp_js_1.Logger.log({ data: result.contentType, level: sp_pnp_js_1.LogLevel.Info, message: `Content Type added successfully.` });
                resolve();
            }, reject);
        });
    }
}
exports.Lists = Lists;
