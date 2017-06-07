import { HandlerBase } from "./handlerbase";
import { IList } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Lists Object Handler
 */
export declare class Lists extends HandlerBase {
    private lists;
    private tokenRegex;
    /**
     * Creates a new instance of the Lists class
     */
    constructor();
    /**
     * Provisioning lists
     *
     * @param lists The lists to provision
     */
    ProvisionObjects(web: Web, lists: IList[]): Promise<void>;
    /**
     * Processes a list
     *
     * @param web The web
     * @param list The list
     */
    private processList(web, conf);
    /**
    * Processes security for a list
    *
    * @param conf The list configuration
    * @param list The pnp list
    */
    private processSecurity(conf, list);
    /**
     * Processes security for a list
     *
     * @param list The pnp list
     * @param roleAssignment Role assignment
     */
    private processRoleAssignment(list, roleAssignment);
    /**
     * Processes content type bindings for a list
     *
     * @param conf The list configuration
     * @param list The pnp list
     * @param contentTypeBindings Content type bindings
     * @param removeExisting Remove existing content type bindings
     */
    private processContentTypeBindings(conf, list, contentTypeBindings, removeExisting);
    /**
     * Processes a content type binding for a list
     *
     * @param conf The list configuration
     * @param list The pnp list
     * @param contentTypeID The Content Type ID
     */
    private processContentTypeBinding(conf, list, contentTypeID);
    /**
     * Processes fields for a list
     *
     * @param web The web
     * @param list The pnp list
     */
    private processFields(web, list);
    /**
     * Processes a field for a lit
     *
     * @param web The web
     * @param conf The list configuration
     * @param fieldXml Field xml
     */
    private processField(web, conf, fieldXml);
    /**
   * Processes field refs for a list
   *
   * @param web The web
   * @param list The pnp list
   */
    private processFieldRefs(web, list);
    /**
     * Processes a field ref for a list
     *
     * @param web The web
     * @param conf The list configuration
     * @param fieldRef The list field ref
     */
    private processFieldRef(web, conf, fieldRef);
    /**
     * Processes views for a list
     *
     * @param web The web
     * @param conf The view configuration
     */
    private processViews(web, conf);
    /**
     * Processes a view for a list
     *
     * @param web The web
     * @param conf List configuration
     * @param view The view configuration
     */
    private processView(web, conf, view);
    /**
     * Processes view fields for a view
     *
     * @param view The pnp view
     * @param viewFields Array of view fields
     */
    private processViewFields(view, viewFields);
    /**
     * Replaces tokens in field xml
     *
     * @param fieldXml The field xml
     */
    private replaceFieldXmlTokens(fieldXml);
}
