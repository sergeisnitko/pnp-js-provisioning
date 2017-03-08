import { HandlerBase } from "./handlerbase";
import { IList } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Features Object Handler
 */
export declare class Lists extends HandlerBase {
    private lists;
    private tokenRegex;
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor();
    /**
     * Provisioning lists
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web: Web, lists: IList[]): Promise<void>;
    private processList(web, list);
    private processContentTypeBindings(list, contentTypeBindings);
    private processContentTypeBinding(list, contentTypeBinding);
    private processFields(web, list);
    private processField(web, list, fieldXml);
    private processViews(list, views);
    private processView(list, view);
    private replaceFieldXmlTokens(fieldXml);
}
