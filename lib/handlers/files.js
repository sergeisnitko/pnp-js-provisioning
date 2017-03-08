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
            if (typeof window === "undefined") {
                reject("Files Handler not supported in Node.");
            }
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
    /**
     * Procceses a file
     *
     * @param web The web
     * @param file The file
     * @param serverRelativeUrl ServerRelativeUrl for the web
     */
    processFile(web, file, serverRelativeUrl) {
        return new Promise((resolve, reject) => {
            fetch(file.Src, { credentials: "include", method: "GET" }).then(res => {
                res.text().then(responseText => {
                    let blob = new Blob([responseText], {
                        type: "text/plain",
                    });
                    let folderServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("", serverRelativeUrl, file.Folder);
                    web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.Url, blob, file.Overwrite).then(result => {
                        Promise.all([
                            this.processWebParts(file, serverRelativeUrl, result.data.ServerRelativeUrl),
                            this.processProperties(web, result, file.Properties),
                        ]).then(resolve, reject);
                    }, reject);
                });
            });
        });
    }
    /**
     * Processes web parts
     *
     * @param file The file
     * @param webServerRelativeUrl ServerRelativeUrl for the web
     * @param fileServerRelativeUrl ServerRelativeUrl for the file
     */
    processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            (new Promise((_resolve, _reject) => {
                if (file.RemoveExistingWebParts) {
                    let ctx = new SP.ClientContext(webServerRelativeUrl), spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared), webParts = lwpm.get_webParts();
                    ctx.load(webParts);
                    ctx.executeQueryAsync(() => {
                        webParts.get_data().forEach(wp => wp.deleteWebPart());
                        ctx.executeQueryAsync(_resolve, _reject);
                    }, _reject);
                }
                else {
                    _resolve();
                }
            })).then(() => {
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
            }, reject);
        });
    }
    /**
     * Process list item properties for the file
     *
     * @param web The web
     * @param result The file add result
     * @param properties The properties to set
     */
    processProperties(web, result, properties) {
        return new Promise((resolve, reject) => {
            if (properties && Object.keys(properties).length > 0) {
                result.file.listItemAllFields.select("ID", "ParentList/ID").expand("ParentList").get().then(({ ID, ParentList }) => {
                    web.lists.getById(ParentList.Id).items.getById(ID).update(properties).then(resolve, reject);
                }, reject);
            }
            else {
                resolve();
            }
        });
    }
    /**
     * Replaces tokens in a string, e.g. {site}
     *
     * @param str The string
     * @param ctx Client context
     */
    replaceTokens(str, ctx) {
        return str
            .replace(/{site}/, sp_pnp_js_1.Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url()));
    }
}
exports.Files = Files;
