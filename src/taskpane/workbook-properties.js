/* global Excel, DOMParser, XMLSerializer */
const propertiesNS = 'http://schemas.microsoft.com/office/2006/metadata/properties';
const contentTypeNS = 'http://schemas.microsoft.com/office/2006/metadata/contentType';
const attributesNS = 'http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes';
const partnerNS = 'http://schemas.microsoft.com/office/infopath/2007/PartnerControls';
const xsiNS = 'http://www.w3.org/2001/XMLSchema-instance';
const xsdNS = 'http://www.w3.org/2001/XMLSchema';
function parse(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Workbook server properties contain invalid XML.');
    return doc;
}
export function rewriteServerProperties(propertiesXml, schemaXml, fields, people = {}, contentType = null) {
    const doc = parse(propertiesXml), schema = parse(schemaXml);
    if (doc.documentElement.namespaceURI !== propertiesNS || schema.documentElement.namespaceURI !== contentTypeNS) {
        throw new Error('Unexpected workbook server property schema.');
    }
    const management = Array.from(doc.getElementsByTagName('*')).find(node => node.localName === 'documentManagement');
    if (!management) throw new Error('Workbook server property container is missing.');
    const definitions = Array.from(schema.getElementsByTagNameNS(xsdNS, 'element'));
    const byField = new Map(fields.map(field => [field.FieldName, field.FieldValue]));
    for (const definition of definitions) {
        const fieldName = definition.getAttributeNS(attributesNS, 'taxonomyFieldName') || definition.getAttributeNS(attributesNS, 'internalName');
        if (!byField.has(fieldName) || definition.getAttributeNS(attributesNS, 'readOnly') === 'true') continue;
        const namespace = definition.parentNode.getAttribute('targetNamespace');
        if (!namespace) continue;
        const name = definition.getAttribute('name');
        let element = Array.from(management.childNodes).find(node => node.nodeType === 1 && node.localName === name && node.namespaceURI === namespace);
        if (!element) { element = doc.createElementNS(namespace, name); management.appendChild(element); }
        while (element.firstChild) element.removeChild(element.firstChild);
        element.removeAttributeNS(xsiNS, 'nil');
        const value = String(byField.get(fieldName) || '');
        const append = (parent, ns, tag, text) => {
            const child = doc.createElementNS(ns, tag); if (text !== undefined) child.textContent = text;
            parent.appendChild(child); return child;
        };
        if (definition.getAttributeNS(attributesNS, 'taxonomy') === 'true') {
            const terms = append(element, partnerNS, 'pc:Terms');
            for (const pair of value ? value.split(';') : []) {
                const separator = pair.lastIndexOf('|');
                const id = pair.substring(separator + 1);
                if (separator < 1 || !/^[0-9a-f]{8}-[0-9a-f-]{27}$/i.test(id)) throw new Error('Workbook taxonomy contains an invalid term identifier.');
                const term = append(terms, partnerNS, 'pc:TermInfo');
                append(term, partnerNS, 'pc:TermName', pair.substring(0, separator));
                append(term, partnerNS, 'pc:TermId', id);
            }
        } else if (definition.getAttributeNS(attributesNS, 'list') === 'UserInfo') {
            const accounts = value ? JSON.parse(value) : [];
            for (const account of accounts) {
                const user = append(element, namespace, 'UserInfo');
                append(user, namespace, 'DisplayName', people[fieldName] || account.Key);
                append(user, namespace, 'AccountId', account.Key);
                append(user, namespace, 'AccountType', 'User');
            }
        } else if (fieldName === 'Audience') {
            for (const choice of value ? value.split(';#') : []) append(element, namespace, 'Value', choice);
        } else if (definition.getAttributeNS(attributesNS, 'format') === 'DateOnly' && value) {
            const match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(value);
            if (!match) throw new Error('Workbook review date format is invalid.');
            element.textContent = `${match[3]}-${match[1].padStart(2, '0')}-${match[2].padStart(2, '0')}T00:00:00`;
        } else {
            element.textContent = value;
            if (!value && definition.getAttribute('nillable') === 'true') element.setAttributeNS(xsiNS, 'xsi:nil', 'true');
        }
    }
    if (contentType) {
        schema.documentElement.setAttributeNS(attributesNS, 'ma:contentTypeID', contentType.id);
        schema.documentElement.setAttributeNS(attributesNS, 'ma:contentTypeName', contentType.name);
    }
    return {properties: new XMLSerializer().serializeToString(doc), schema: new XMLSerializer().serializeToString(schema)};
}
export async function syncWorkbookServerProperties(fields, columns, people, contentType, excel = Excel) {
    return excel.run(async context => {
        const parts = context.workbook.customXmlParts;
        parts.load('items');
        const custom = context.workbook.properties.custom;
        custom.load('items/key');
        await context.sync();
        // Inspect root namespaces directly: downloaded SharePoint parts may use
        // prefixed roots and do not always advertise namespace references.
        const xmlParts = parts.items.map(part => ({part, xml: part.getXml()}));
        await context.sync();
        const properties = xmlParts.filter(item => parse(item.xml.value).documentElement.namespaceURI === propertiesNS);
        const schemas = xmlParts.filter(item => parse(item.xml.value).documentElement.namespaceURI === contentTypeNS);
        if (properties.length > 1 || schemas.length > 1 || properties.length !== schemas.length) {
            throw new Error('Workbook has ambiguous SharePoint server properties. Metadata save stopped.');
        }
        if (properties.length) {
            const updated = rewriteServerProperties(properties[0].xml.value, schemas[0].xml.value, fields, people, contentType);
            properties[0].part.setXml(updated.properties);
            if (contentType) schemas[0].part.setXml(updated.schema);
        }
        // Server metadata lives in documentManagement XML. Remove duplicate custom
        // properties for these fields so stale aliases cannot override that XML.
        const managed = new Set(fields.map(field => field.FieldName.toLowerCase()));
        for (const column of Object.values(columns)) {
            if (managed.has(column.name.toLowerCase())) managed.add(column.displayName.toLowerCase());
        }
        for (const property of custom.items) if (managed.has(property.key.toLowerCase()) && property.key.toLowerCase() !== 'contenttypeid') property.delete();
        if (contentType) custom.add('ContentTypeId', contentType.id);
        const title = fields.find(field => field.FieldName === 'Title');
        if (title) context.workbook.properties.title = title.FieldValue;
        await context.sync();
    });
}
