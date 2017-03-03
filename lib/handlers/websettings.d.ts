import { HandlerBase } from "./handlerbase";
import { IWebSettings } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Features Object Handler
 */
export declare class WebSettings extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor();
    /**
     * Provisioning features
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web: Web, settings: IWebSettings): Promise<void>;
}
