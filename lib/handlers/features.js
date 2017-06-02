"use strict";
const handlerbase_1 = require("./handlerbase");
/**
 * Describes the Features Object Handler
 */
class Features extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Features");
    }
    /**
     * Provisioning features
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web, features) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            features.reduce((chain, feature) => {
                if (feature.deactivate) {
                    return chain.then(() => web.features.remove(feature.id, feature.force));
                }
                else {
                    return chain.then(() => web.features.add(feature.id, feature.force));
                }
            }, Promise.resolve({})).then(() => {
                super.scope_ended();
                resolve();
            }).catch((e) => {
                super.scope_ended();
                reject(e);
            });
        });
    }
}
exports.Features = Features;
