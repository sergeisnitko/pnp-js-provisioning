import { IComposedLook } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web } from "sp-pnp-js";
import { ReplaceTokens } from "../util";

/**
 * Describes the Composed Look Object Handler
 */
export class ComposedLook extends HandlerBase {
    /**
     * Creates a new instance of the ObjectComposedLook class
     */
    constructor() {
        super("ComposedLook");
    }

    /**
     * Provisioning Composed Look
     * 
     * @param object The Composed Look to provision
     */
    public ProvisionObjects(web: Web, composedLook: IComposedLook): Promise<void> {
        super.scope_started();
        return new Promise<void>((resolve, reject) => {
            web.applyTheme(
                ReplaceTokens(composedLook.ColorPaletteUrl),
                ReplaceTokens(composedLook.FontSchemeUrl),
                composedLook.BackgroundImageUrl ? ReplaceTokens(composedLook.BackgroundImageUrl) : null,
                false).then(_ => {
                    super.scope_ended();
                    resolve();
                }).catch(e => {
                    super.scope_ended();
                    reject(e);
                });
        });
    }
}

