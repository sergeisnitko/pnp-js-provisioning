// we need to import HandlerBase & TypedHash to avoid naming issues in ts transpile
import { Schema } from "./schema";
import { HandlerBase } from "./handlers/handlerbase";
import { TypedHash, Web, Logger, LogLevel } from "sp-pnp-js";
import { DefaultHandlerMap } from "./handlers/exports";

/**
 * Root class of Provisioning 
 */
export class WebProvisioner {

    /**
     * Creates a new instance of the Provisioner class
     * 
     * @param web The Web instance to which we want to apply templates
     * @param handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
     */
    constructor(private web: Web, public handlerMap: TypedHash<HandlerBase> = DefaultHandlerMap) { }

    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     * 
     * @param template The template to apply
     */
    public applyTemplate(template: Schema): Promise<void> {

        Logger.write(`Beginning processing of web [${this.web.toUrl()}]`, LogLevel.Info);

        return Object.getOwnPropertyNames(template).sort((name1: string, name2: string) => {

            if (name1 === name2) {

                return 0;
            }

            return 0;


            // this needs to be more complex and control the order of the elements
            // things like site columns should be created before they may need added to a list, etc.
            // so a sort function 

            //  if (a is less than b by some ordering criterion) {
            //     return -1;
            //   }
            //   if (a is greater than b by the ordering criterion) {
            //     return 1;
            //   }
            //   // a must be equal to b
            //   return 0;

        }).reduce((chain, name) => {

            let handler = this.handlerMap[name];

            return chain.then(_ => handler.ProvisionObjects(this.web, template[name]));

        }, Promise.resolve()).then(_ => {

            Logger.write(`Done processing of web [${this.web.toUrl()}]`, LogLevel.Info);
        });
    }
}
