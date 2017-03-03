import { HandlerBase } from "./handlerbase";
import { IFile } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Features Object Handler
 */
export declare class Files extends HandlerBase {
    /**
     * Creates a new instance of the Files class
     */
    constructor();
    /**
     * Provisioning Files
     *
     * @paramm files The files  to provision
     */
    ProvisionObjects(web: Web, files: IFile[]): Promise<void>;
    private processFile(web, file, serverRelativeUrl);
}
