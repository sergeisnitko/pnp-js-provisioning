"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Features Object Handler
 */
class Files extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Files class
     */
    constructor() {
        super("Files");
    }
    /**
     * Provisioning Files
     *
     * @paramm files The files  to provision
     */
    ProvisionObjects(web, files) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            web.get().then(({ ServerRelativeUrl }) => {
                files.reduce((chain, file) => chain.then(_ => this.processFile(web, file, ServerRelativeUrl)), Promise.resolve()).then(() => {
                    super.scope_ended();
                    resolve();
                }).catch(e => {
                    super.scope_ended();
                    reject(e);
                });
            });
        });
    }
    processFile(web, file, serverRelativeUrl) {
        return new Promise((resolve, reject) => {
            if (Blob) {
                fetch(file.Src, { credentials: "include", method: "GET" }).then(res => {
                    res.text().then(responseText => {
                        let blob = new Blob([responseText], {
                            type: "text/plain",
                        });
                        let folderServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("", serverRelativeUrl, file.Folder);
                        web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.Url, blob, file.Overwrite).then(resolve, reject);
                    });
                });
            }
            else {
                reject();
            }
        });
    }
}
exports.Files = Files;
