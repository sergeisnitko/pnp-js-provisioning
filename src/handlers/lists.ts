import { HandlerBase } from "./handlerbase";
import { IList, IContentTypeBinding } from "../schema";
import { Web, List, Logger, LogLevel } from "sp-pnp-js";

/**
 * Describes the Features Object Handler
 */
export class Lists extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Lists");
    }

    /**
     * Provisioning lists
     * 
     * @paramm features The features to provision
     */
    public ProvisionObjects(web: Web, lists: IList[]): Promise<void> {
        super.scope_started();
        return new Promise<void>((resolve, reject) => {
            lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve()).then(() => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }

    private processList(web: Web, list: IList): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            web.lists.ensure(list.Title, list.Description, list.Template, list.ContentTypesEnabled, list.AdditionalSettings).then(result => {
                if (result.created) {
                    Logger.log({ data: result.list, level: LogLevel.Info, message: `List ${list.Title} created successfully.` });
                }
                this.processContentTypeBindings(result.list, list.ContentTypeBindings).then(resolve, reject);
            });
        });
    }

    private processContentTypeBindings(list: List, contentTypeBindings: IContentTypeBinding[]): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            if (contentTypeBindings) {
                contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(list, ct)), Promise.resolve()).then(() => {
                    super.scope_ended();
                    resolve();
                }).catch(e => {
                    super.scope_ended();
                    reject(e);
                });
            } else {
                resolve();
            }
        });
    }

    private processContentTypeBinding(list: List, contentTypeBinding: IContentTypeBinding): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            list.contentTypes.addAvailableContentType(contentTypeBinding.ContentTypeID).then(result => {
                Logger.log({ data: result.contentType, level: LogLevel.Info, message: `Content Type added successfully.` });
                resolve();
            }, reject);
        });
    }
}
