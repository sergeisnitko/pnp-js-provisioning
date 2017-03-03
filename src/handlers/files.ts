import { HandlerBase } from "./handlerbase";
import { IFile } from "../schema";
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
                    console.log(res);
                    let blob = new Blob(["This is my blob content"], {
                        type: "text/plain",
                    });
                    let folderServerRelativeUrl = Util.combinePaths("", serverRelativeUrl, file.Folder);
                    web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.Url, blob, file.Overwrite).then(resolve, reject);
                });
            } else {
                reject();
            }
        });
    }
}
