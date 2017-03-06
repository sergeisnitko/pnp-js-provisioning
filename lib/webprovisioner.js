"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const sp_pnp_js_1 = require("sp-pnp-js");
const exports_1 = require("./handlers/exports");
/**
 * Root class of Provisioning
 */
class WebProvisioner {
    /**
     * Creates a new instance of the Provisioner class
     *
     * @param web The Web instance to which we want to apply templates
     * @param handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
     */
    constructor(web, handlerMap = exports_1.DefaultHandlerMap, handlerSort = exports_1.DefaultHandlerSort) {
        this.web = web;
        this.handlerMap = handlerMap;
        this.handlerSort = handlerSort;
    }
    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     *
     * @param template The template to apply
     */
    applyTemplate(template) {
        sp_pnp_js_1.Logger.write(`Beginning processing of web [${this.web.toUrl()}]`, sp_pnp_js_1.LogLevel.Info);
        // keeping this broken out allows for easier debugging of the incoming tasks + ordering
        let operations = Object.getOwnPropertyNames(template).sort((name1, name2) => {
            let sort1 = this.handlerSort.hasOwnProperty(name1) ? this.handlerSort[name1] : 99;
            let sort2 = this.handlerSort.hasOwnProperty(name2) ? this.handlerSort[name2] : 99;
            return sort1 - sort2;
        });
        console.log(operations);
        // reduce those operations to a promise chain and return that. When this chain resolves the site is provisioned
        return operations.reduce((chain, name) => {
            let handler = this.handlerMap[name];
            return chain.then(_ => handler.ProvisionObjects(this.web, template[name]));
        }, Promise.resolve()).then(_ => {
            sp_pnp_js_1.Logger.write(`Done processing of web [${this.web.toUrl()}]`, sp_pnp_js_1.LogLevel.Info);
        });
    }
}
exports.WebProvisioner = WebProvisioner;
