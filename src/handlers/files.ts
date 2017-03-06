import { HandlerBase } from "./handlerbase";
import { IFile, IWebPart } from "../schema";
import { Web, Util } from "sp-pnp-js";

/**
 * Describes the Features Object Handler
 */
export class Files extends HandlerBase {
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
    public ProvisionObjects(web: Web, files: IFile[]): Promise<void> {

        super.scope_started();

        return new Promise<void>((resolve, reject) => {
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

    private processFile(web: Web, file: IFile, serverRelativeUrl: string): Promise<any> {
        return new Promise((resolve, reject) => {
            if (Blob) {
                fetch(file.Src, { credentials: "include", method: "GET" }).then(res => {
                    res.text().then(responseText => {
                        let blob = new Blob([responseText], {
                            type: "text/plain",
                        });
                        let folderServerRelativeUrl = Util.combinePaths("", serverRelativeUrl, file.Folder);
                        web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.Url, blob, file.Overwrite).then(({ data }) => {
                            this.processWebParts(file, serverRelativeUrl, data.ServerRelativeUrl).then(resolve, reject);
                        }, reject);
                    });
                });
            } else {
                reject();
            }
        });
    }

    private processWebParts(file: IFile, webServerRelativeUrl: string, fileServerRelativeUrl: string) {
        return new Promise((resolve, reject) => {
            if (file.WebParts && file.WebParts.length > 0) {
                let ctx = new SP.ClientContext(webServerRelativeUrl),
                    spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl),
                    lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                file.WebParts.forEach(wp => {
                    let def = lwpm.importWebPart(wp.Contents.Xml),
                        inst = def.get_webPart();
                    lwpm.addWebPart(inst, wp.Zone, wp.Order);
                    ctx.load(inst);
                });
                ctx.executeQueryAsync(resolve, reject);
            } else {
                resolve();
            }
        });
    }
}
