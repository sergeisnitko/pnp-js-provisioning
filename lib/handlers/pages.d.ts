import { HandlerBase } from "./handlerbase";
import { INavigation } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Features Object Handler
 */
export declare class Navigation extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor();
    /**
     * Provisioning features
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web: Web, navigation: INavigation): Promise<void>;
    private processNavTree(target, nodes);
    private processNode(target, node);
    private deleteExistingNodes(target);
    private deleteNode(target, id);
}
