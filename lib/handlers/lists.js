"use strict";
const xmljs = require("xml-js");
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Lists Object Handler
 */
class Lists extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Lists class
     */
    constructor() {
        super("Lists");
        this.tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g;
        this.lists = [];
    }
    /**
     * Provisioning lists
     *
     * @param lists The lists to provision
     */
    ProvisionObjects(web, lists) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve()).then(() => {
                lists.reduce((chain, list) => chain.then(_ => this.processFields(web, list)), Promise.resolve()).then(() => {
                    lists.reduce((chain, list) => chain.then(_ => this.processFieldRefs(web, list)), Promise.resolve()).then(() => {
                        lists.reduce((chain, list) => chain.then(_ => this.processViews(web, list)), Promise.resolve()).then(() => {
                            super.scope_ended();
                            resolve();
                        });
                    });
                });
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
    /**
     * Processes a list
     *
     * @param web The web
     * @param list The list
     */
    processList(web, conf) {
        return new Promise((resolve, reject) => {
            web.lists.ensure(conf.Title, conf.Description, conf.Template, conf.ContentTypesEnabled, conf.AdditionalSettings).then(({ created, list, data }) => {
                this.lists.push(data);
                if (created) {
                    sp_pnp_js_1.Logger.log({ data: list, level: sp_pnp_js_1.LogLevel.Info, message: `List ${conf.Title} created successfully.` });
                }
                this.processContentTypeBindings(conf, list, conf.ContentTypeBindings, conf.RemoveExistingContentTypes).then(resolve, reject);
            });
        });
    }
    /**
     * Processes content type bindings for a list
     *
     * @param conf The list configuration
     * @param list The pnp list
     * @param contentTypeBindings Content type bindings
     * @param removeExisting Remove existing content type bindings
     */
    processContentTypeBindings(conf, list, contentTypeBindings, removeExisting) {
        return new Promise((resolve, reject) => {
            if (contentTypeBindings) {
                contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(conf, list, ct.ContentTypeID)), Promise.resolve()).then(() => {
                    if (removeExisting) {
                        let promises = [];
                        list.contentTypes.get().then(contentTypes => {
                            contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                                let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                                    && (ContentTypeId.indexOf("0x0120") === -1);
                                if (shouldRemove) {
                                    sp_pnp_js_1.Logger.write(`Removing content type ${ContentTypeId} from list ${conf.Title}`, sp_pnp_js_1.LogLevel.Info);
                                    promises.push(list.contentTypes.getById(ContentTypeId).delete());
                                }
                            });
                        });
                        Promise.all(promises).then(resolve, reject);
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
    /**
     * Processes a content type binding for a list
     *
     * @param conf The list configuration
     * @param list The pnp list
     * @param contentTypeID The Content Type ID
     */
    processContentTypeBinding(conf, list, contentTypeID) {
        return new Promise((resolve, reject) => {
            list.contentTypes.addAvailableContentType(contentTypeID).then(({ contentType }) => {
                sp_pnp_js_1.Logger.log({ data: contentType, level: sp_pnp_js_1.LogLevel.Info, message: `Content Type ${contentTypeID} added successfully to list ${conf.Title}.` });
                resolve();
            }, reject);
        });
    }
    /**
     * Processes fields for a list
     *
     * @param web The web
     * @param list The pnp list
     */
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
    /**
     * Processes a field for a lit
     *
     * @param web The web
     * @param conf The list configuration
     * @param fieldXml Field xml
     */
    processField(web, conf, fieldXml) {
        return new Promise((resolve, reject) => {
            let fieldProps = JSON.parse(xmljs.xml2json(fieldXml)), { InternalName, DisplayName } = fieldProps.elements[0].attributes;
            fieldProps.elements[0].attributes.DisplayName = InternalName;
            web.lists.getByTitle(conf.Title).fields.createFieldAsXml(this.replaceFieldXmlTokens(xmljs.json2xml(fieldProps))).then(({ data, field }) => {
                field.update({ Title: DisplayName }).then(() => {
                    sp_pnp_js_1.Logger.log({ data: data, level: sp_pnp_js_1.LogLevel.Info, message: `Field '${DisplayName}' added successfully to list ${conf.Title}.` });
                    resolve();
                }, reject);
            });
        });
    }
    /**
   * Processes field refs for a list
   *
   * @param web The web
   * @param list The pnp list
   */
    processFieldRefs(web, list) {
        return new Promise((resolve, reject) => {
            if (list.FieldRefs) {
                list.FieldRefs.reduce((chain, fieldRef) => chain.then(_ => this.processFieldRef(web, list, fieldRef)), Promise.resolve()).then(resolve, reject);
            }
            else {
                resolve();
            }
        });
    }
    /**
     * Processes a field ref for a list
     *
     * @param web The web
     * @param conf The list configuration
     * @param fieldRef The list field ref
     */
    processFieldRef(web, conf, fieldRef) {
        return new Promise((resolve, reject) => {
            web.lists.getByTitle(conf.Title).fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName }).then(() => {
                sp_pnp_js_1.Logger.log({ data: fieldRef, level: sp_pnp_js_1.LogLevel.Info, message: `Field '${fieldRef.ID}' updated for list ${conf.Title}.` });
                resolve();
            }, reject);
        });
    }
    /**
     * Processes views for a list
     *
     * @param web The web
     * @param conf The view configuration
     */
    processViews(web, conf) {
        return new Promise((resolve, reject) => {
            if (conf.Views) {
                conf.Views.reduce((chain, view) => chain.then(_ => this.processView(web, conf, view)), Promise.resolve()).then(resolve, reject);
            }
            else {
                resolve();
            }
        });
    }
    /**
     * Processes a view for a list
     *
     * @param web The web
     * @param conf List configuration
     * @param view The view configuration
     */
    processView(web, conf, view) {
        return new Promise((resolve, reject) => {
            let _view = web.lists.getByTitle(conf.Title).views.getByTitle(view.Title);
            _view.get().then(_ => {
                _view.update(view.AdditionalSettings).then(() => {
                    this.processViewFields(_view, view.ViewFields).then(resolve, reject);
                });
            }, () => {
                web.lists.getByTitle(conf.Title).views.add(view.Title, view.PersonalView, view.AdditionalSettings).then(result => {
                    sp_pnp_js_1.Logger.log({ data: result.data, level: sp_pnp_js_1.LogLevel.Info, message: `View ${view.Title} added successfully to list ${conf.Title}.` });
                    this.processViewFields(result.view, view.ViewFields).then(resolve, reject);
                }, reject);
            });
        });
    }
    /**
     * Processes view fields for a view
     *
     * @param view The pnp view
     * @param viewFields Array of view fields
     */
    processViewFields(view, viewFields) {
        return new Promise((resolve, reject) => {
            view.fields.removeAll().then(() => {
                viewFields.reduce((chain, viewField) => chain.then(_ => view.fields.add(viewField)), Promise.resolve()).then(resolve, reject);
            }, reject);
        });
    }
    /**
     * Replaces tokens in field xml
     *
     * @param fieldXml The field xml
     */
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
