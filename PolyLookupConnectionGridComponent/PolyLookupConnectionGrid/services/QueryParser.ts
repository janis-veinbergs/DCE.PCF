import { IEntityDefinition, IMetadata } from "PolyLookupConnectionGrid/types/metadata";
import { createElement } from "react"

//Credit to Scott Durrow: https://github.com/scottdurow/SparkleXrm/blob/master/SparkleXrmSamples/ConnectionsUI/ClientUI/ViewModel/QueryParser.cs

/**
 * a createElement overload which offers a way to create elements, their attributes as objects and also their children. 
 * @param tagName - Element tag which to create
 * @param attributes - Attributes for element
 * @param children - If string, that would be a text node for tagName element, otherwise another element.
 * @returns 
 */
export function createElementAttributes(tagName: string, attributes?: {[key: string]: string}, ...children: (Element | string)[]) : Element {
    attributes = attributes || {};
    const nativeOptions = !!attributes.is ? { is: attributes.is } : undefined;
    delete attributes.is;
    const element = document.createElementNS("", tagName, nativeOptions);
    Object.entries(attributes).forEach(([name,value]) => name.startsWith('on') ? 
        (element as Record<string, any>)[name] = value : element.setAttribute(name,value)
    );
    children
      .filter( child => !(child == null || child == undefined))
      .forEach( child => element.appendChild( 
        child instanceof Node ? child : document.createTextNode(child)
      ));
    return element;
  };

function xmlEncodeForFetchXml(strInput: string): string {
    ///<summary>Sanitizes strInput to be safely used in FetchXML queries</summary>
    return strInput.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
  
  
function parseXml(input: string) {
    return new DOMParser().parseFromString(input, "application/xml");
}

export type SearchWildcards = "none" | "prefixWildcard" | "suffixWildcard" | "bothWildcard"

type QueryParserSettings = {
    entities: string[],
    quickFindCount: number
}

type FetchQuerySettings = {
    displayName?: string | null,
    columns: LayoutJsonRow[],
    fetchXml: string | null,
    rootEntity: EntityLookup | null,
    orderByAttribute?: string | null,
    orderByDesending?: boolean | null
}


export type EntityLookup = {
    logicalName: string,
    aliasName?: string,
    views: {[viewName: string]: FetchQuerySettings},
    // attributes: {[key: string]: AttributeQuery},
    quickFindQuery?: FetchQuerySettings,
    metadata: IMetadata
}

type SavedQueryFetchXmlResult = {
    /** Use the query part in the URL in the nextLink property as the value for the options parameter in your subsequent retrieveMultipleRecords call to request the next set of records. Don't change or append any more system query options to the value. For every subsequent request for more pages, you should use the same maxPageSize value used in the original retrieve multiple request. Also, cache the results returned or the value of the nextLink property so that previously retrieved pages can be returned.
     * The value of the nextLink property is URI encoded. If you URI encode the value before you send it, the XML cookie information in the URL will cause an error. */
    nextLink: string | undefined,
    entities: SavedQueryEntity[]
}
type SavedQueryEntity = {
    '@odata.etag': string,
    fetchxml: string,
    isquickfindquery: boolean,
    layoutjson: string,
    layoutxml: string,
    name: string,
    returnedtypecode: string,
    savedqueryid: string
}

type LayoutJsonRow = {
    AddedBy: string,
    CellType: string,
    Desc: string,
    DisableMetaDataBinding: boolean,
    DisableSorting: boolean,
    ImageProviderFunctionName: string,
    ImageProviderWebresource: string,
    IsHidden: boolean,
    LabelId: string,
    Name: string,
    RelatedEntityName: string,
    Width: number
}

type LayoutJson = {
    CustomControlDescriptions: unknown[],
    Icon: boolean,
    IconRenderer: string,
    Jump: string,
    Name: string,
    Object: number,
    Preview: boolean,
    Rows: LayoutJsonRow[],
    Select: boolean
}
//#endregion


function getFetchXmlParentFilter(query: FetchQuerySettings, parentAttribute: string): string {
    if (!query.fetchXml) { throw "fetchXml missing from query"; }
    const fetchXmlDom = parseXml(query.fetchXml);
    const fetchElement = fetchXmlDom.querySelector('fetch');
    if (fetchElement === null) {throw "fetchXml from query doesn't contain <fetch> tag"}
    fetchElement.setAttribute('count', '{0}');
    fetchElement.setAttribute('paging-cookie', '{1}');
    fetchElement.setAttribute('page', '{2}');
    fetchElement.setAttribute('returntotalrecordcount', 'true');
    fetchElement.setAttribute('distinct', 'true');
    fetchElement.setAttribute('no-lock', 'true');
    const orderByElement = fetchElement.querySelector('order');
    if (orderByElement != null) {
        query.orderByAttribute = orderByElement.getAttribute('attribute');
        query.orderByDesending = orderByElement.getAttribute('descending') === 'true';
        orderByElement.remove();
    }
    let filter = fetchElement.querySelector('entity>filter');
    if (filter !== null) {
        const filterType = filter.getAttribute('type');
        if (filterType === 'or') {
            const andFilter = createElementAttributes('filter',
                {
                    'type': 'and'
                },
                filter.innerHTML
            );
            filter.remove();
            filter = andFilter;
            fetchElement.querySelector('entity')!.append(andFilter);
        }
        const parentFilter = createElementAttributes('condition', {
            'attribute': parentAttribute,
            'operator': 'eq',
            'value': `"#ParentRecordPlaceholder#"`
        })
        filter.append(parentFilter);
    }
    
    return fetchXmlDom.documentElement.outerHTML.replaceAll('</entity>', '{3}</entity>');
}

// export function getQuickFind(entityLogicalName: string) {
//     return _getViewDefinition(entityLogicalName, true, null);
// };

// export function getView(entityLogicalName: string, viewName: string) {
//     return _getViewDefinition(entityLogicalName, false, viewName);
// };

// function _getViewDefinition(entityLogicalName: string, isQuickFind: boolean, viewName: string | null) {
//     const metadataPromise = WebApi.invokeRetrieveEntityDefinitionsPromise<EntityLookup["metadata"]>(
//         entityLogicalName,
//         "?$select=PrimaryIdAttribute,PrimaryNameAttribute,ObjectTypeCode,EntitySetName,DisplayName,LogicalName,IconVectorName" + 
//         "&$expand=Attributes(" + 
//             "$select=AttributeOf,AttributeType,LogicalName;" + 
//             "$filter=AttributeOf ne null or AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'Lookup'" + 
//         ")"
//     );
//     return metadataPromise
//         .then(entityMetadata => {
//             const metadata = entityMetadata;
//             //Takes current fetchXml, appends filter, passes result to next invocation.
//             let fetchXml = 
//                 `<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
//                     <entity name='savedquery'>
//                     <attribute name='name' />
//                     <attribute name='fetchxml' />
//                     <attribute name='layoutjson' />
//                     <attribute name='returnedtypecode' />
//                     <filter type='and'>
//                         <condition attribute="statecode" operator="eq" value="0" />
//                         <condition attribute="querytype" operator="in">
//                             <value>0</value>
//                             <value>4</value>
//                             <value>8192</value>
//                         </condition>
//                         <filter type='or'>
//                             <condition attribute='returnedtypecode' operator='eq' value='${metadata.ObjectTypeCode.toString()}'/>
//                         </filter>
//                 ${  isQuickFind ?
//                         "<condition attribute='isquickfindquery' operator='eq' value='1'/><condition attribute='isdefault' operator='eq' value='1'/>"
//                     : viewName != null && viewName.length > 0 ?
//                         `<condition attribute='name' operator='eq' value='${Utils.xmlEncodeForFetchXml(viewName)}'/>"`
//                     :
//                         "<condition attribute='querytype' operator='eq' value='2'/><condition attribute='isdefault' operator='eq' value='1'/>"
//                 }
//                     </filter>
//                 </entity>
//                 </fetch>`

//             const savedQueriesPromise = Xrm.WebApi.retrieveMultipleRecords("savedquery", `?fetchXml=${fetchXml.trim().replaceAll(/[\r\n]/g, "")}`) as Promise<SavedQueryFetchXmlResult>
//             return savedQueriesPromise
//                 .then(savedQueries => {
//                     if (savedQueries.entities.length === 0) {
//                         throw "_getViewDefinition yielded 0 savedViews";
//                     }
//                     let savedQuery = savedQueries.entities[0]
//                     let query: EntityLookup = {
//                             logicalName: savedQuery.returnedtypecode,
//                             views: {},
//                             metadata: metadata
//                         };
//                     let config = _parse(savedQuery.fetchxml, JSON.parse(savedQuery.layoutjson) as LayoutJson, metadata);
//                     query.views[savedQuery.name] = config;
//                     if (isQuickFind) {
//                         query.quickFindQuery = config;
//                     }
//                     return query;
//                 });
//         })
        
// };

enum QuickFindPlaceholder {
    Text_0 = '{0}',
    Int_1 = '{1}',
    Currency_2 = '{2}',
    DateTime_3 = '{3}',
    Float_4 = '{4}',
}

function _parse(fetchXml: string, layoutJson: LayoutJson, metadata: EntityLookup["metadata"], quickFindCount?: number) {
    const querySettings: FetchQuerySettings = {
        fetchXml: fetchXml,
        rootEntity: _parseFetchXml(parseXml(fetchXml), metadata),
        columns: layoutJson.Rows
    };
    return querySettings;
};

function _parseFetchXml(fetchXmlDom: Document, metadata: EntityLookup["metadata"]) {
    const entityElement = fetchXmlDom.querySelector('entity')!;
    const logicalName = entityElement.getAttribute('name')!;
    
    const rootEntity : EntityLookup = {
        logicalName: logicalName,
        // attributes: {},
        views: {},
        metadata: metadata
    };
    // const linkEntities = entityElement.querySelectorAll('link-entity');
    // linkEntities.forEach(element => {
    //     const link: EntityLookup = {
    //         // attributes: {},
    //         aliasName: element.getAttribute('alias') ?? undefined,
    //         logicalName: element.getAttribute('name')!,
    //         views: {}
    //     }
        // if (!(link.logicalName in this.entityLookup)) {
        //     this.entityLookup[link.logicalName] = link;
        // }
        // else {
        //     const alias = link.aliasName;
        //     link = this.entityLookup[link.logicalName];
        //     link.aliasName = alias;
        // }
        // if (link.aliasName && !(link.aliasName in this._aliasEntityLookup)) {
        //     this._aliasEntityLookup[link.aliasName] = link;
        // }
    // });
    return rootEntity;
    // const conditions = fetchElement.querySelectorAll("filter[isquickfindfields='1'] > condition");
    // conditions.forEach((element, index) => {
    //     let logicalName = element.getAttribute('attribute')!;
    //     if (!(logicalName in rootEntity.attributes)) {
    //         let attribute: AttributeQuery = {
    //             logicalName: logicalName,
    //             columns: new Array()
    //         };
    //         rootEntity.attributes[logicalName] = attribute;
    //     }
    // });
};


type SearchOptions = {
    /** Whether to prefix or suffix (or both) with wildcard. */
    wildcards?: SearchWildcards,
    /** Impose count on fetch or not */
    recordLimit?: number,
    /** add no-lock to fetch or not. By default adds. */
    nolock?: boolean,
    /** Do not preserve order defined within fetchxml */
    removeOrder?: boolean,
    /** By default yes */
    distinct?: boolean
}
/**
 * Return FetchXml that can be executed against CRM where placeholders are replaced with searchTerm. Issues distinct query.
 * If searchTerm is NOT parseable to either int/currency/datetime/float, those conditions are removed.
 * @param config - get with getQuickFind or getView.
 * @param searchTerm - whatever you search for.
 * @param attributeMetadata - pass AttributeMetadata (only 2 properties needed). QuickFind must have knowledge of lookup properties to have them searchable. https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/attributemetadata?view=dataverse-latest
 * @param options - Adjust how things are searched
 * @param additionalFilter - add custom filter in addition to the filter already present in the view to be used for searching. Pass <filter type="and/or"> tag
 * @param additionalAttributes - any additional attributes you want to include within FetchXml. Only single <entity> element within FetchXml supported. Only primary entity attributes, linked entity attributes cannot be added. Use additionalLinkAttributes to add additional <link-entity> with respective attributes
 * @param additionalLinkAttributes - Add additional link-entity. Useful if want to fetch additional attribute from linked entity. to/from entities can be duplicate/alreayd existing under entity element and all specified attributes will be fetched.
 *                                Example: ['<link-entity name="account" from="accountid" to="deac_accountid" visible="false" link-type="outer" alias="accountnamealias"><attribute name="name" /></link-entity>']
 * @returns 
 */
export function getFetchXmlForQuery(
    fetchXml: string,
    searchTerm?: string,
    attributeMetadata?: IEntityDefinition["Attributes"],
    { wildcards = "none", recordLimit, nolock = true, removeOrder = false, distinct = true}: Partial<SearchOptions> = {},
    additionalFilter?: string,
    additionalAttributes?: string[],
    additionalLinkAttributes?: string[]
) : string
{
    const fetchElement = parseXml(fetchXml).documentElement;
    if (fetchElement.tagName !== 'fetch') { throw "fetchElement expected to be <fetch> tag" }
    fetchElement.setAttribute('distinct', distinct ? 'true' : 'false');
    fetchElement.setAttribute('no-lock', nolock ? 'true' : 'false');
    if (recordLimit) {
        fetchElement?.setAttribute('count', recordLimit.toString());
    }
    const orderByElement = fetchElement.querySelector('order');
    removeOrder && orderByElement?.remove();

    //For lookup attributes, if we directly filter by lookup attribute, it expects guid and will return error: An exception System.FormatException was thrown while trying to convert input value 'a%' to attribute 'deac_servicecontract.deac_accountid'. Expected type of attribute value: System.Guid. Exception raised: Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).
    //So we must replace it to extended virtual attribute that has suffix: name
    const doNotQueryDirectlyAttributes = attributeMetadata?.filter(x => x.AttributeOf);
    if (additionalFilter && additionalFilter != "") {
        addCustomFilter(fetchElement, additionalFilter);
    }
    if (additionalAttributes && additionalAttributes.length > 0) {
        addAdditionalAttributes(fetchElement, additionalAttributes);
    }
    if (additionalLinkAttributes && additionalLinkAttributes.length > 0) {
        addAdditionalLinkAttributes(fetchElement, additionalLinkAttributes);
    }

    const conditions = fetchElement.querySelectorAll("filter[isquickfindfields='1'] > condition");
    conditions.forEach(element => {
        const logicalName = element.getAttribute('attribute');
        const found = doNotQueryDirectlyAttributes?.find(x => x.AttributeOf === logicalName);
        if (found) {
            element.setAttribute('attribute', found.LogicalName);
        }
    });
    
    //QuickFind fetchXml contains placeholders where searchable value should go into.
    searchTerm && !isNaN(parseInt(searchTerm))
        ? fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.Int_1}']`)?.forEach(x => x.setAttribute("value", parseInt(searchTerm).toString()))
        : fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.Int_1}']`)?.forEach(x => x.remove())
    searchTerm && !isNaN(parseFloat(searchTerm))
        ? fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.Currency_2}']`)?.forEach(x => x.setAttribute("value", parseFloat(searchTerm).toString()))
        : fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.Currency_2}']`)?.forEach(x => x.remove());
    searchTerm && !isNaN(Date.parse(searchTerm))
        ? fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.DateTime_3}']`)?.forEach(x => x.setAttribute("value", new Date(Date.parse(searchTerm)).toLocaleDateString('en-US')))
        : fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.DateTime_3}']`)?.forEach(x => x.remove());
    searchTerm && !isNaN(parseFloat(searchTerm))
        ? fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.Float_4}"']`)?.forEach(x => x.setAttribute("value", parseFloat(searchTerm).toString()))
        : fetchElement.querySelectorAll(`condition[value='${QuickFindPlaceholder.Float_4}"']`)?.forEach(x => x.remove());


    if (searchTerm !== undefined) {
        let textSearchTerm = searchTerm ?? "";
        if (wildcards == "prefixWildcard" || wildcards == "bothWildcard") {
            textSearchTerm = '%' + textSearchTerm.substring(textSearchTerm.startsWith('*') ? 1 : 0);
        }
        if (wildcards == "suffixWildcard" || wildcards == "bothWildcard") {
            textSearchTerm = textSearchTerm.substring(0, textSearchTerm.length - (textSearchTerm.endsWith('*') ? 1 : 0)) + '%';
        }
        const fetchXmlModified = fetchElement.outerHTML.replaceAll(QuickFindPlaceholder.Text_0, xmlEncodeForFetchXml(textSearchTerm))
        return fetchXmlModified;
    }
    return fetchElement.outerHTML;
}

function addCustomFilter(fetchXml: HTMLElement, customFilter: string) {
    if (fetchXml.tagName !== "fetch") { throw new Error("fetch root tag expected"); }
    //Ensure we have parseAble customFilter
    const customFilterElement = parseXml(customFilter).documentElement;
    const firstFilter = fetchXml.querySelector("filter") as Element;
    if (firstFilter) {
        //Wrap custom filter like this:
        //Given <filter type="or"><condition.... /></filter>
        //Do: <filter type="and"><filter type="or"><condition.... /></filter>${customFilter}</filter>
        const wrapper = createElementAttributes("filter", { type: "and" });
        firstFilter.parentElement!.insertBefore(wrapper, firstFilter);
        wrapper.appendChild(firstFilter);
        wrapper.appendChild(customFilterElement);

        //const mergedFilter = createElementAttributes("filter", { type: "and" }, firstFilter.cloneNode(true) as typeof firstFilter, customFilterElement);
        //firstFilter.remove();
        //fetchXml.querySelector("fetch")?.appendChild(mergedFilter);
    } else {
        //No current filter exists - add filter element as is
        fetchXml.querySelector("entity")!.appendChild(customFilterElement);
    }
}

function addAdditionalAttributes(fetchXml: HTMLElement, additionalAttributes: string[]) {
    if (addAdditionalAttributes.length === 0) { return; }
    if (fetchXml.tagName !== "fetch") { throw new Error("fetch root tag expected"); }
    const entityElements = fetchXml.querySelectorAll("entity");
    if (entityElements.length == 0) { throw new Error("No entity element specified, cannot addAdditionalAttributes"); }
    if (entityElements.length > 2) { throw new Error("addAdditionalAttributes supported only for single <entity> element"); }
    const entityElement = entityElements[0];
    if (entityElement.querySelector("all-attributes") !== null) {
        //No additional attributes need to be added when all being fetched
        return;
    }
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    const existingAttributeNames = Array.from(entityElement.querySelectorAll("attribute")).map(x => x.getAttribute("name"))!;
    additionalAttributes
        .filter(x => !existingAttributeNames.includes(x))
        .forEach(x => entityElement.appendChild(
            createElementAttributes("attribute", {name: x}))
        );
}

function addAdditionalLinkAttributes(fetchXml: HTMLElement, additionalLinkAttributes: string[]) {
    if (addAdditionalAttributes.length === 0) { return; }
    if (fetchXml.tagName !== "fetch") { throw new Error("fetch root tag expected"); }
    const entityElements = fetchXml.querySelectorAll("entity");
    if (entityElements.length == 0) { throw new Error("No entity element specified, cannot addAdditionalLinkAttributes"); }
    if (entityElements.length > 2) { throw new Error("addAdditionalAttributes supported only for single <entity> element"); }
    const entityElement = entityElements[0];
    additionalLinkAttributes.forEach(x => entityElement.appendChild(parseXml(x).documentElement));
}