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
                        web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.Url, blob, file.Overwrite).then(({ data }) => {
                            this.processWebParts(file, serverRelativeUrl, data.ServerRelativeUrl).then(resolve, reject);
                        }, reject);
                    });
                });
            }
            else {
                reject();
            }
        });
    }
    processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            if (file.WebParts && file.WebParts.length > 0) {
                let ctx = new SP.ClientContext(webServerRelativeUrl), spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                file.WebParts.forEach(wp => {
                    let def = lwpm.importWebPart(this.replaceTokens(wp.Contents.Xml, ctx)), inst = def.get_webPart();
                    lwpm.addWebPart(inst, wp.Zone, wp.Order);
                    ctx.load(inst);
                });
                ctx.executeQueryAsync(resolve, reject);
            }
            else {
                resolve();
            }
        });
    }
    replaceTokens(str, ctx) {
        return str
            .replace(/{site}/, sp_pnp_js_1.Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url()));
    }
}
exports.Files = Files;
