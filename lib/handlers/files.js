"use strict";
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
const util_1 = require("../util");
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
            fetch(util_1.ReplaceTokens(file.Src), { credentials: "include", method: "GET" }).then(res => {
                res.text().then(responseText => {
                    let blob = new Blob([responseText], {
                        type: "text/plain",
                    });
                    let folderServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("", serverRelativeUrl, file.Folder);
                    web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.Url, blob, file.Overwrite).then(result => {
                        Promise.all([
                            this.processWebParts(file, serverRelativeUrl, result.data.ServerRelativeUrl),
                            this.processProperties(web, result, file.Properties),
                        ]).then(_ => {
                            this.processPageListViews(web, file.WebParts, result.data.ServerRelativeUrl).then(resolve, reject);
                        }, reject);
                    }, reject);
                });
            });
        });
    }
    /**
     * Remove exisiting webparts if specified
     *
     * @param webServerRelativeUrl ServerRelativeUrl for the web
     * @param fileServerRelativeUrl ServerRelativeUrl for the file
     * @param shouldRemove Should web parts be removed
     */
    removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, shouldRemove) {
        return new Promise((resolve, reject) => {
            if (shouldRemove) {
                let ctx = new SP.ClientContext(webServerRelativeUrl), spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared), webParts = lwpm.get_webParts();
                ctx.load(webParts);
                ctx.executeQueryAsync(() => {
                    webParts.get_data().forEach(wp => wp.deleteWebPart());
                    ctx.executeQueryAsync(resolve, reject);
                }, reject);
            }
            else {
                resolve();
            }
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
            this.removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, file.RemoveExistingWebParts).then(() => {
                if (file.WebParts && file.WebParts.length > 0) {
                    let ctx = new SP.ClientContext(webServerRelativeUrl), spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                    file.WebParts.forEach(wp => {
                        let def = lwpm.importWebPart(this.replaceXmlTokens(wp.Contents.Xml, ctx)), inst = def.get_webPart();
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
    processPageListViews(web, webParts, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            if (webParts) {
                let listViewWebParts = webParts.filter(wp => wp.ListView);
                if (listViewWebParts.length > 0) {
                    listViewWebParts
                        .reduce((chain, wp) => chain.then(_ => this.processPageListView(web, wp.ListView, fileServerRelativeUrl)), Promise.resolve())
                        .then(resolve, reject);
                }
                else {
                    resolve();
                }
            }
            else {
                resolve();
            }
        });
    }
    processPageListView(web, listView, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            let views = web.lists.getByTitle(listView.List).views;
            views.get().then(listViews => {
                let wpView = listViews.filter(v => v.ServerRelativeUrl === fileServerRelativeUrl);
                if (wpView.length === 1) {
                    let view = views.getById(wpView[0].Id);
                    let settings = listView.View.AdditionalSettings || {};
                    view.update(settings).then(() => {
                        view.fields.removeAll().then(_ => {
                            listView.View.ViewFields.reduce((chain, viewField) => chain.then(() => view.fields.add(viewField)), Promise.resolve()).then(resolve, reject);
                        }, reject);
                    }, reject);
                }
                else {
                    resolve();
                }
            });
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
    replaceXmlTokens(str, ctx) {
        let site = sp_pnp_js_1.Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url());
        return str.replace(/{site}/g, site);
    }
}
exports.Files = Files;
