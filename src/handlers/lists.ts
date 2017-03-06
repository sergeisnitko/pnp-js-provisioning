import { HandlerBase } from "./handlerbase";
import { IList, IContentTypeBinding } from "../schema";
import { Web, List, Logger, LogLevel } from "sp-pnp-js";

/**
 * Describes the Features Object Handler
 */
export class Lists extends HandlerBase {
    private lists: any[];
    private tokenRegex = /{[a-z]*:[A-za-z]*}/g;
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Lists");
        this.lists = [];
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
                lists.reduce((chain, list) => chain.then(_ => this.processFields(web, list)), Promise.resolve()).then(() => {
                    super.scope_ended();
                    resolve();
                });
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }

    private processList(web: Web, list: IList): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            web.lists.ensure(list.Title, list.Description, list.Template, list.ContentTypesEnabled, list.AdditionalSettings).then(result => {
                this.lists.push(result.data);
                if (result.created) {
                    Logger.log({ data: result.list, level: LogLevel.Info, message: `List ${list.Title} created successfully.` });
                }
                Promise.all([
                    this.processContentTypeBindings(result.list, list.ContentTypeBindings)
                ]).then(() => {
                    resolve();
                }, reject);
            });
        });
    }

    private processContentTypeBindings(list: List, contentTypeBindings: IContentTypeBinding[]): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            if (contentTypeBindings) {
                contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(list, ct)), Promise.resolve()).then(() => {
                    resolve();
                }).catch(e => {
                    reject(e);
                });
            } else {
                resolve();
            }
        });
    }

    private processContentTypeBinding(list: List, contentTypeBinding: IContentTypeBinding): Promise<any> {
        return new Promise<void>((resolve, reject) => {
            list.contentTypes.addAvailableContentType(contentTypeBinding.ContentTypeID).then(result => {
                Logger.log({ data: result.contentType, level: LogLevel.Info, message: `Content Type added successfully.` });
                resolve();
            }, reject);
        });
    }

    private processFields(web: Web, list: IList): Promise<any> {
        return new Promise<void>((resolve, reject) => {
            if (list.Fields) {
                list.Fields.reduce((chain, field) => chain.then(_ => this.processField(web, list, field)), Promise.resolve()).then(resolve, reject);
            } else {
                resolve();
            }
        });
    }

    private processField(web: Web, list: IList, fieldXml: string): Promise<any> {
        return web.lists.getByTitle(list.Title).fields.createFieldAsXml(this.replaceFieldXmlTokens(fieldXml));
    }

    private replaceFieldXmlTokens(fieldXml: string) {
        let m;
        while ((m = this.tokenRegex.exec(fieldXml)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach((match) => {
                let [Type, Value] = match.replace(/[\{\}]/g, "").split(":");
                switch (Type) {
                    case "listid": {
                        let list = this.lists.filter(l => l.Title === Value);
                        if (list.length === 1) {
                            fieldXml = fieldXml.replace(match, list[0].Id);
                        }
                    }
                }
            });
        }
        return fieldXml;
    }
}
