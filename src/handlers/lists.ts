import { HandlerBase } from "./handlerbase";
import { IList } from "../schema";
import { Web, Logger, LogLevel } from "sp-pnp-js";

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

        return web.lists.ensure(list.Title, list.Description, list.Template, list.ContentTypesEnabled, list.AdditionalSettings).then(result => {

            if (result.created) {
                Logger.log({ data: result.list, level: LogLevel.Info, message: `List ${list.Title} created successfully.` });
            }

            // here we would do things like add fields, apply content types, etc
        });
    }
}
