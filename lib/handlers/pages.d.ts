import { HandlerBase } from "./handlerbase";
import { IPage } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Pages Object Handler
 */
export declare class Pages extends HandlerBase {
    /**
     * Creates a new instance of the ObjectPages class
     */
    constructor();
    /**
     * Provisioning pages
     *
     * @paramm pages The pages to provision
     */
    ProvisionObjects(web: Web, pages: IPage[]): Promise<void>;
    private processPage(web, page, serverRelativeUrl);
}
