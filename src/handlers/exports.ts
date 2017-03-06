import { TypedHash } from "sp-pnp-js";
import { HandlerBase } from "./handlerbase";
import { ComposedLook } from "./composedlook";
import { CustomActions } from "./customactions";
import { Features } from "./features";
import { WebSettings } from "./websettings";
import { Navigation } from "./navigation";
import { Lists } from "./lists";
import { Files } from "./files";

export const DefaultHandlerMap: TypedHash<HandlerBase> = {
    ComposedLook: new ComposedLook(),
    CustomActions: new CustomActions(),
    Features: new Features(),
    Lists: new Lists(),
    Navigation: new Navigation(),
    WebSettings: new WebSettings(),
    Files: new Files(),
};

export const DefaultHandlerSort: TypedHash<number> = {

    WebSettings: 1,
    Features: 2,
    Files: 4,
    Lists: 3,
    Navigation: 7,
    CustomActions: 5,
    ComposedLook: 6,
}
