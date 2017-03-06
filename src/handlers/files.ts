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

    private processFile(web: Web, file: IFile, serverRelativeUrl: string): Promise<any> {
        return new Promise((resolve, reject) => {
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
        });
    }

    private processWebParts(file: IFile, webServerRelativeUrl: string, fileServerRelativeUrl: string) {
        return new Promise((resolve, reject) => {
            (new Promise((_resolve, _reject) => {
                if (file.RemoveExistingWebParts) {
                    let ctx = new SP.ClientContext(webServerRelativeUrl),
                        spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl),
                        lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared),
                        webParts = lwpm.get_webParts();
                    ctx.load(webParts);
                    ctx.executeQueryAsync(() => {
                        webParts.get_data().forEach(wp => wp.deleteWebPart());
                        ctx.executeQueryAsync(_resolve, _reject);
                    }, _reject);
                } else {
                    _resolve();
                }
            })).then(() => {
                if (file.WebParts && file.WebParts.length > 0) {
                    let ctx = new SP.ClientContext(webServerRelativeUrl),
                        spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl),
                        lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                    file.WebParts.forEach(wp => {
                        let def = lwpm.importWebPart(this.replaceTokens(wp.Contents.Xml, ctx)),
                            inst = def.get_webPart();
                        lwpm.addWebPart(inst, wp.Zone, wp.Order);
                        ctx.load(inst);
                    });
                    ctx.executeQueryAsync(resolve, reject);
                } else {
                    resolve();
                }
            }, reject);
        });
    }

    private replaceTokens(str: string, ctx: SP.ClientContext): string {
        return str
            .replace(/{site}/, Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url()));
    }
}
