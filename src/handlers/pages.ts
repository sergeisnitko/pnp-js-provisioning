import { HandlerBase } from "./handlerbase";
import { IPage } from "../schema";
import { Web, Util } from "sp-pnp-js";

/**
 * Describes the Pages Object Handler
 */
export class Pages extends HandlerBase {
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
    public ProvisionObjects(web: Web, pages: IPage[]): Promise<void> {
        super.scope_started();

        return new Promise<void>((resolve, reject) => {
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

    private processPage(web: Web, page: IPage, serverRelativeUrl: string): Promise<void> {
        let folderServerRelativeUrl = Util.combinePaths("", serverRelativeUrl, page.Folder),
            fileUrl = Util.combinePaths("", folderServerRelativeUrl, page.Url);
        return new Promise<void>((resolve, reject) => {
            web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.addTemplateFile(fileUrl, 1).then(({ file }) => {
                if (page.Fields) {
                    file.listItemAllFields.select("ID", "ParentList/ID").expand("ParentList").get().then(({ ID, ParentList }) => {
                        web.lists.getById(ParentList.Id).items.getById(ID).update(page.Fields).then(() => {
                            resolve();
                        }, reject);
                    }, reject);
                } else {
                    resolve();
                }
            });
        });
    }
}
