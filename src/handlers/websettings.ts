import { HandlerBase } from "./handlerbase";
import { IWebSettings } from "../schema";
import { Web } from "sp-pnp-js";
import { ReplaceTokens } from "../util";

/**
 * Describes the Features Object Handler
 */
export class WebSettings extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("WebSettings");
    }

    /**
     * Provisioning features
     * 
     * @paramm features The features to provision
     */
    public ProvisionObjects(web: Web, settings: IWebSettings): Promise<void> {
        super.scope_started();
        return new Promise<void>((resolve, reject) => {
            Object.keys(settings)
                .filter(key => typeof (settings[key]) === "string")
                .forEach(key => {
                    let value: string = <any>settings[key];
                    settings[key] = ReplaceTokens(value);
                });
            web.update(settings).then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => reject(e));
        });
    }
}
