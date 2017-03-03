"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Features Object Handler
 */
class Navigation extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Navigation");
    }
    /**
     * Provisioning features
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web, navigation) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            let chain = Promise.resolve();
            if (sp_pnp_js_1.Util.isArray(navigation.QuickLaunch)) {
                chain.then(_ => this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
            }
            if (sp_pnp_js_1.Util.isArray(navigation.TopNavigationBar)) {
                chain.then(_ => this.processNavTree(web.navigation.topNavigationBar, navigation.TopNavigationBar));
            }
            return chain.then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
    processNavTree(target, nodes) {
        return nodes.reduce((chain, node) => chain.then(_ => this.processNode(target, node)), Promise.resolve());
    }
    processNode(target, node) {
        return target.add(node.Title, node.Url).then(result => {
            if (sp_pnp_js_1.Util.isArray(node.Children)) {
                return this.processNavTree(result.node.children, node.Children);
            }
        });
    }
}
exports.Navigation = Navigation;
