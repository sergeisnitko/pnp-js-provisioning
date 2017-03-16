"use strict";
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Features Object Handler
 */
class Lists extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Lists");
        this.tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z]*}/g;
        this.lists = [];
    }
    /**
     * Provisioning lists
     *
     * @paramm features The features to provision
     */
    ProvisionObjects(web, lists) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve()).then(() => {
                lists.reduce((chain, list) => chain.then(_ => this.processFields(web, list)), Promise.resolve()).then(() => {
                    lists.reduce((chain, list) => chain.then(_ => this.processViews(web, list)), Promise.resolve()).then(() => {
                        super.scope_ended();
                        resolve();
                    });
                });
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
    processList(web, list) {
        return new Promise((resolve, reject) => {
            web.lists.ensure(list.Title, list.Description, list.Template, list.ContentTypesEnabled, list.AdditionalSettings).then(result => {
                this.lists.push(result.data);
                if (result.created) {
                    sp_pnp_js_1.Logger.log({ data: result.list, level: sp_pnp_js_1.LogLevel.Info, message: `List ${list.Title} created successfully.` });
                }
                this.processContentTypeBindings(result.list, list.ContentTypeBindings, list.RemoveExistingContentTypes).then(resolve, reject);
            });
        });
    }
    processContentTypeBindings(list, contentTypeBindings, removeExisting) {
        return new Promise((resolve, reject) => {
            if (contentTypeBindings) {
                contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(list, ct)), Promise.resolve()).then(() => {
                    if (removeExisting) {
                        if (typeof window === "undefined") {
                            reject("Removal of existing content types not supported in Node.");
                        }
                        else {
                            let promises = [];
                            list.contentTypes.get().then(contentTypes => {
                                contentTypes.forEach(({ Id: { StringValue } }) => {
                                    let shouldRemove = (contentTypeBindings.filter(ctb => StringValue.indexOf(ctb.ContentTypeID) === -1).length > 0);
                                    if (shouldRemove) {
                                        promises.push(list.contentTypes.getById(StringValue).delete());
                                    }
                                });
                            });
                            Promise.all(promises).then(resolve, reject);
                        }
                    }
                    else {
                        resolve();
                    }
                }).catch(e => {
                    reject(e);
                });
            }
            else {
                resolve();
            }
        });
    }
    processContentTypeBinding(list, contentTypeBinding) {
        return new Promise((resolve, reject) => {
            list.contentTypes.addAvailableContentType(contentTypeBinding.ContentTypeID).then(result => {
                sp_pnp_js_1.Logger.log({ data: result.contentType, level: sp_pnp_js_1.LogLevel.Info, message: `Content Type added successfully.` });
                resolve();
            }, reject);
        });
    }
    processFields(web, list) {
        return new Promise((resolve, reject) => {
            if (list.Fields) {
                list.Fields.reduce((chain, field) => chain.then(_ => this.processField(web, list, field)), Promise.resolve()).then(resolve, reject);
            }
            else {
                resolve();
            }
        });
    }
    processField(web, list, fieldXml) {
        return web.lists.getByTitle(list.Title).fields.createFieldAsXml(this.replaceFieldXmlTokens(fieldXml));
    }
    processViews(web, list) {
        return new Promise((resolve, reject) => {
            if (list.Views) {
                list.Views.reduce((chain, view) => chain.then(_ => this.processView(web, list, view)), Promise.resolve()).then(resolve, reject);
            }
            else {
                resolve();
            }
        });
    }
    processView(web, list, view) {
        return new Promise((resolve, reject) => {
            let _view = web.lists.getByTitle(list.Title).views.getByTitle(view.Title);
            _view.get().then(_ => {
                this.processViewFields(_view, view.ViewFields).then(resolve, reject);
            }, () => {
                web.lists.getByTitle(list.Title).views.add(view.Title, view.PersonalView, view.AdditionalSettings).then(result => {
                    this.processViewFields(result.view, view.ViewFields).then(resolve, reject);
                }, reject);
            });
        });
    }
    processViewFields(view, viewFields) {
        return new Promise((resolve, reject) => {
            view.fields.removeAll().then(() => {
                viewFields.reduce((chain, viewField) => chain.then(_ => view.fields.add(viewField)), Promise.resolve()).then(resolve, reject);
            }, reject);
        });
    }
    replaceFieldXmlTokens(fieldXml) {
        let m;
        while ((m = this.tokenRegex.exec(fieldXml)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach((match) => {
                let [Type, Value] = match.replace(/[\{\}]/g, "").split(":");
                switch (Type) {
                    case "listid": {
                        let list = this.lists.filter(l => l.Title === Value);
                        if (list.length === 1) {
                            fieldXml = fieldXml.replace(match, list[0].Id);
                        }
                    }
                }
            });
        }
        return fieldXml;
    }
}
exports.Lists = Lists;
