"use strict";
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Pages Object Handler
 */
class Pages extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectPages class
     */
    constructor() {
        super("Pages");
    }
    /**
     * Provisioning pages
     *
     * @paramm pages The pages to provision
     */
    ProvisionObjects(web, pages) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            web.get().then(({ ServerRelativeUrl }) => {
                pages.reduce((chain, page) => chain.then(_ => this.processPage(web, page, ServerRelativeUrl)), Promise.resolve()).then(() => {
                    super.scope_ended();
                    resolve();
                }).catch(e => {
                    super.scope_ended();
                    reject(e);
                });
            });
        });
    }
    processPage(web, page, serverRelativeUrl) {
        let folderServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("", serverRelativeUrl, page.Folder), fileUrl = sp_pnp_js_1.Util.combinePaths("", folderServerRelativeUrl, page.Url);
        return new Promise((resolve, reject) => {
            web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.addTemplateFile(fileUrl, 1).then(({ file }) => {
                if (page.Fields) {
                    file.listItemAllFields.select("ID", "ParentList/ID").expand("ParentList").get().then(({ ID, ParentList }) => {
                        web.lists.getById(ParentList.Id).items.getById(ID).update(page.Fields).then(() => {
                            resolve();
                        }, reject);
                    }, reject);
                }
                else {
                    resolve();
                }
            });
        });
    }
}
exports.Pages = Pages;
