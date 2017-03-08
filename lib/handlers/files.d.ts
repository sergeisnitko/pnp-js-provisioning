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
    /**
     * Procceses a file
     *
     * @param web The web
     * @param file The file
     * @param serverRelativeUrl ServerRelativeUrl for the web
     */
    private processFile(web, file, serverRelativeUrl);
    /**
     * Processes web parts
     *
     * @param file The file
     * @param webServerRelativeUrl ServerRelativeUrl for the web
     * @param fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl);
    /**
     * Process list item properties for the file
     *
     * @param web The web
     * @param result The file add result
     * @param properties The properties to set
     */
    private processProperties(web, result, properties);
    /**
     * Replaces tokens in a string, e.g. {site}
     *
     * @param str The string
     * @param ctx Client context
     */
    private replaceTokens(str, ctx);
}
