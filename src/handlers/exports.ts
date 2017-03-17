import { TypedHash } from "sp-pnp-js";
import { HandlerBase } from "./handlerbase";
import { ComposedLook } from "./composedlook";
import { CustomActions } from "./customactions";
import { Features } from "./features";
import { WebSettings } from "./websettings";
import { Navigation } from "./navigation";
import { Lists } from "./lists";
import { Pages } from "./pages";

export const DefaultHandlerMap: TypedHash<HandlerBase> = {
    ComposedLook: new ComposedLook(),
    CustomActions: new CustomActions(),
    Features: new Features(),
    Lists: new Lists(),
    Navigation: new Navigation(),
    Pages: new Pages(),
    WebSettings: new WebSettings(),
};

export const DefaultHandlerSort: TypedHash<number> = {
    ComposedLook: 6,
    CustomActions: 5,
    Features: 2,
    Lists: 3,
    Navigation: 4,
    Pages: 7,
    WebSettings: 1,
};
