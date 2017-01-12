import { HandlerBase } from "./handlerbase";
import { INavigation, INavigationNode } from "../schema";
import { Web, NavigationNodes, Util } from "sp-pnp-js";

/**
 * Describes the Features Object Handler
 */
export class Navigation extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Navigation");
    }

    /**
     * Provisioning features
     * 
     * @paramm features The features to provision
     */
    public ProvisionObjects(web: Web, navigation: INavigation): Promise<void> {

        super.scope_started();

        return new Promise<void>((resolve, reject) => {

            let chain = Promise.resolve();

            if (Util.isArray(navigation.QuickLaunch)) {
                chain.then(_ => this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
            }

            if (Util.isArray(navigation.TopNavigationBar)) {
                chain.then(_ => this.processNavTree(web.navigation.topNavigationBar, navigation.TopNavigationBar));
            }

            return chain.then(_ => {

                super.scope_ended();
                resolve();

            }).catch(e => {

                super.scope_ended();
                reject(e);
            });
        });
    }

    private processNavTree(target: NavigationNodes, nodes: INavigationNode[]): Promise<void> {

        return nodes.reduce((chain, node) => chain.then(_ => this.processNode(target, node)), Promise.resolve());
    }

    private processNode(target: NavigationNodes, node: INavigationNode): Promise<void> {

        return target.add(node.Title, node.Url).then(result => {

            if (Util.isArray(node.Children)) {
                return this.processNavTree(result.node.children, node.Children);
            }
        });
    }
}
