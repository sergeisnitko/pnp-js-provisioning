import { HandlerBase } from "./handlerbase";
import { IPropertyBagEntry } from "../schema";
import { Web } from "sp-pnp-js";

/**
 * Describes the PropertyBagEntries Object Handler
 */
export class PropertyBagEntries extends HandlerBase {
    /**
     * Creates a new instance of the PropertyBagEntries class
     */
    constructor() {
        super("PropertyBagEntries");
    }

    /**
     * Provisioning property bag entries
     * 
     * @paramm entries The property bag entries to provision
     */
    public ProvisionObjects(web: Web, entries: IPropertyBagEntry[]): Promise<void> {
        super.scope_started();
        return new Promise<any>((resolve, reject) => {
            if (typeof window === "undefined") {
                reject("PropertyBagEntries Handler not supported in Node.");
            }
            web.get().then(({ ServerRelativeUrl }) => {
                let ctx = new SP.ClientContext(ServerRelativeUrl),
                    spWeb = ctx.get_web(),
                    propBag = spWeb.get_allProperties(),
                    idxProps = [];
                entries.filter(entry => entry.Overwrite).forEach(entry => {
                    propBag.set_item(entry.Key, entry.Value);
                    if (entry.Indexed) {
                        idxProps.push(this.EncodePropertyKey(entry.Key));
                    }
                });
                spWeb.update();
                ctx.load(propBag);
                ctx.executeQueryAsync(() => {
                    if (idxProps.length > 0) {
                        propBag.set_item("vti_indexedpropertykeys", idxProps.join("|"));
                        spWeb.update();
                        ctx.executeQueryAsync(() => {
                            super.scope_ended();
                            resolve();
                        }, () => {
                            super.scope_ended();
                            reject();
                        });
                    } else {
                        super.scope_ended();
                        resolve();
                    }
                }, () => {
                    super.scope_ended();
                    reject();
                });
            });
        });
    }

    /**
     *Encode property key
     * 
     * @param propKey Property bag key
     */
    private EncodePropertyKey(propKey: string): string {
        let bytes = [];
        for (let i = 0; i < propKey.length; ++i) {
            bytes.push(propKey.charCodeAt(i));
            bytes.push(0);
        }
        let b64encoded = window.btoa(String.fromCharCode.apply(null, bytes));
        return b64encoded;
    }
}
