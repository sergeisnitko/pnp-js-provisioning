import { HandlerBase } from "./handlerbase";
import { IFeature } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Features Object Handler
 */
export declare class Features extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor();
    /**
     * Provisioning features
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web: Web, features: IFeature[]): Promise<void>;
}
