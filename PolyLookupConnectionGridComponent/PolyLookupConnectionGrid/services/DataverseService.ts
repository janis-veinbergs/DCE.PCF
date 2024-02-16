import { useQueries, useQuery, UseQueryResult } from "@tanstack/react-query";
import axios from "axios";
import { LanguagePack } from "../types/languagePack";
import {
  IEntityDefinition,
  IManyToManyRelationship,
  IMetadata,
  IOneToManyRelationship,
  isManyToMany,
  IViewDefinition,
  IViewLayout,
} from "../types/metadata";
import Handlebars from "handlebars";
import { ILookupItem } from "PolyLookupConnectionGrid/components/LookupItem";
import { createElementAttributes, getFetchXmlForQuery } from "./QueryParser";

const nToNColumns = [
  "SchemaName",
  "Entity1LogicalName",
  "Entity2LogicalName",
  "Entity1IntersectAttribute",
  "Entity2IntersectAttribute",
  "RelationshipType",
  "IntersectEntityName",
];

const oneToNColumns = [
  "SchemaName",
  "ReferencedEntity",
  "ReferencingEntity",
  "ReferencingEntityNavigationPropertyName",
  "ReferencedAttribute",
  "ReferencingAttribute",
  "RelationshipType",
];

const tableDefinitionColumns = [
  "LogicalName",
  "DisplayName",
  "PrimaryIdAttribute",
  "PrimaryNameAttribute",
  "EntitySetName",
  "DisplayCollectionName",
  "IsQuickCreateEnabled",
  "IconVectorName"
];

const viewDefinitionColumns = ["savedqueryid", "name", "fetchxml", "layoutjson", "querytype"];

const apiVersion = "9.1";

export function useMetadata(
  currentTable: string,
  relationshipName: string,
  relationship2Name: string | undefined,
  lookupView: string | undefined
) {
  return useQuery({
    queryKey: ["metadata", currentTable, relationshipName, relationship2Name, lookupView],
    queryFn: () => getMetadata(currentTable, relationshipName, relationship2Name, lookupView),
    enabled: !!currentTable && !!relationshipName,
  });
}

export function useMetadataGrid(
  currentTable: string,
  referencedTables: string[],
  relationshipName: string,
  clientUrl: string,
  // lookupView: string | undefined
) {
  console.count(`useMetadataGrid, ${currentTable}, ${referencedTables.join(",")}, ${relationshipName}, ${clientUrl}`);
  const queries = useQueries({
    queries: referencedTables.map((referencedTable) => ({
      queryKey: ["metadata", currentTable, referencedTable, relationshipName, clientUrl],
      queryFn: () => getMetadataGrid(currentTable, referencedTable, relationshipName, clientUrl),
      enabled: !!currentTable && !!referencedTable && !!relationshipName,
    }))
  });
  const metadata: Record<string, UseQueryResult<IMetadata>> = queries.reduce((prev, cur, index) => ({
    ...prev,
    [referencedTables[index]]: cur
  }), {});
  return metadata;
}

export function useSelectedItems(
  metadata: IMetadata | undefined,
  currentRecordId: string,
  formType: XrmEnum.FormType | undefined,
) {
  return useQuery({
    queryKey: ["selectedItems", metadata, currentRecordId, formType],
    queryFn: () =>
      retrieveAssociatedRecords(
        currentRecordId,
        metadata?.intersectEntity.LogicalName,
        metadata?.intersectEntity.EntitySetName,
        metadata?.intersectEntity.PrimaryIdAttribute,
        metadata?.currentIntesectAttribute,
        metadata?.associatedIntesectAttribute,
        metadata?.associatedEntity.LogicalName,
        metadata?.associatedEntity.PrimaryIdAttribute,
        metadata?.associatedEntity.PrimaryNameAttribute,
        metadata?.associatedEntity.IconVectorName,
        metadata?.clientUrl
      ),
    enabled:
      !!metadata?.intersectEntity.EntitySetName &&
      !!metadata?.associatedEntity.EntitySetName &&
      (formType === XrmEnum.FormType.Update ||
        formType === XrmEnum.FormType.ReadOnly ||
        formType === XrmEnum.FormType.Disabled),
  });
}

/** Parse dataset records to ILookupItem */
export function useSelectedItemsGrid(
  currentTable: string,
  relationshipName: string,
  validEntityLogicalNames: string[],
  records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>,
  clientUrl: string,
) {
  const entitiesWithinData = Object.values(records).map(value => {
    const record2id = value.getValue("record2id") as ComponentFramework.EntityReference | null;
      if (!record2id?.etn) {
        throw new Error("record2id 'Record (To)' is mandatory to be used within grid view. Couldn't determine entity logical name");
      }
      return record2id.etn
  });
  const entitiesNeedMedata = entitiesWithinData.concat(validEntityLogicalNames);
  const metadata = useMetadataGrid(currentTable, entitiesNeedMedata, relationshipName, clientUrl);
  return useQuery({
    queryKey: ["useSelectedItemsGrid"].concat(Object.keys(records)),
    queryFn: () => {
      return Object.entries(records)
    .map(([_, value]) => {      
      const record2id = value.getValue("record2id") as ComponentFramework.EntityReference | null;
      if (!record2id?.etn) {
        throw new Error("record2id 'Record (To)' is mandatory to be used within grid view. Couldn't determine entity logical name");
      }
      const entityMetadata = metadata[record2id.etn].data;
      if (!entityMetadata) { throw new Error("entity metadata not ready. Adjust enabled property not to execute query before all metadata isSuccess"); }
      const lookupItem: ILookupItem = {
        entityReference: {
          id: record2id.id.guid,
          name: record2id.name,
          etn: record2id.etn,
        },
        connectionReference: {
          id: value.getRecordId(),
          name: value.getNamedReference().name,
          etn: value.getNamedReference().etn
        },
        entityIconUrl: entityMetadata.associatedEntity.IconVectorName ? `${entityMetadata.clientUrl}/webresources/${entityMetadata.associatedEntity.IconVectorName}` : null,
        key: value.getRecordId(),
        name: record2id.name,
        data: {
          [entityMetadata.associatedEntity.PrimaryIdAttribute as string]: record2id.id.guid,
          [entityMetadata.associatedEntity.PrimaryNameAttribute as string]: record2id.name,
        },
        metadata: entityMetadata
      } as ILookupItem;
      return lookupItem;
    });
    },
    enabled: Object.values(metadata).every(x => x.isSuccess)
  })
}

export function useLanguagePack(webResourcePath: string | undefined, defaultLanguagePack: LanguagePack) {
  return useQuery({
    queryKey: ["languagePack", webResourcePath],
    queryFn: () => getLanguagePack(webResourcePath, defaultLanguagePack),
    enabled: !!webResourcePath,
  });
}

export function getLanguagePack(
  webResourceUrl: string | undefined,
  defaultLanguagePack: LanguagePack
): Promise<LanguagePack> {
  const languagePack: LanguagePack = { ...defaultLanguagePack };

  if (webResourceUrl === undefined) return Promise.resolve(languagePack);

  return axios
    .get(webResourceUrl)
    .then((res) => {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(res.data, "text/xml");

      const nodes = xmlDoc.getElementsByTagName("data");
      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        const key = node.getAttribute("name");
        const value = node.getElementsByTagName("value")[0].childNodes[0].nodeValue ?? "";
        if (key && value) {
          languagePack[key as keyof LanguagePack] = value;
        }
      }

      return languagePack;
    })
    .catch(() => languagePack);
}

export function getManytoManyRelationShipDefinition(
  currentTable: string | undefined,
  relationshipName: string | undefined
) {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<IManyToManyRelationship>(
          `/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${currentTable}')/ManyToManyRelationships(SchemaName='${relationshipName}')`,
          {
            params: {
              $select: nToNColumns.join(","),
            },
          }
        )
        .then((res) => res.data);
}

export function getOnetoManyRelationShipDefinition(
  currentTable: string | undefined,
  relationshipName: string | undefined
) {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<IOneToManyRelationship>(
          `/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${currentTable}')/OneToManyRelationships(SchemaName='${relationshipName}')`,
          {
            params: {
              $select: oneToNColumns.join(","),
            },
          }
        )
        .then((res) => res.data);
}

export function getConnectionRelationShipDefinition(
  referencedEntity: string | undefined
) {
  return typeof referencedEntity === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<{ value: IOneToManyRelationship[] }>(
          `/api/data/v${apiVersion}/RelationshipDefinitions/Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata`,
          {
            params: {
              $select: oneToNColumns.join(","),
              $filter: `ReferencingEntity eq 'connection' and ReferencedEntity eq '${referencedEntity}'`,
            },
          }
        )
        .then((res) => res.data.value.at(1));
}

export function getManytoOneRelationShipDefinition(
  currentTable: string | undefined,
  relationshipName: string | undefined
) {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<{ value: IOneToManyRelationship[] }>(
          `/api/data/v${apiVersion}/RelationshipDefinitions/Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata`,
          {
            params: {
              $select: oneToNColumns.join(","),
              $filter: `SchemaName eq '${relationshipName}' and ReferencingEntity eq '${currentTable}'`,
            },
          }
        )
        .then((res) => res.data.value.at(0));
}

export function getEntityDefinition(entityName: string | undefined) {
  return typeof entityName === "undefined"
    ? Promise.reject(new Error("Invalid entity name"))
    : axios
        .get<IEntityDefinition>(`/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${entityName}')`, {
          params: {
            $select: tableDefinitionColumns.join(","),
            $expand: "Attributes($select=LogicalName,AttributeOf)",
          },
        })
        .then((res) => {
          return {
            ...res.data,
            DisplayNameLocalized: res.data.DisplayName.UserLocalizedLabel?.Label ?? "",
            DisplayCollectionNameLocalized: res.data.DisplayCollectionName.UserLocalizedLabel?.Label ?? "",
          } as IEntityDefinition;
        });
}

export async function getViewDefinition(
  entityName: string | undefined,
  viewName: string | undefined,
  queryType?: number | undefined
) {
  if (typeof entityName === "undefined") return Promise.reject(new Error("Invalid arguments"));

  const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
    params: {
      $filter: `returnedtypecode eq '${entityName}' ${viewName ? `and name eq '${viewName}'` : ""} ${
        queryType ? `and querytype eq ${queryType}` : ""
      } `,
      $select: viewDefinitionColumns.join(","),
    },
  });

  const layoutjson = result.data.value[0].layoutjson as unknown as string;
  const layout = JSON.parse(layoutjson) as IViewLayout;
  result.data.value[0].layoutjson = layout;
  return result.data.value[0];
}


export async function getDefaultView(entityName: string | undefined, viewName: string | undefined = undefined) {
  if (viewName) {
    const viewByName = await getViewDefinition(entityName, viewName);
    if (viewByName) return viewByName;
  }

  const defaultView = await getViewDefinition(entityName, undefined, 64);
  return defaultView;
}

export function useDefaultView(entityName: string | undefined, viewName: string | undefined) {
  return useQuery({
    queryKey: ["defaultView", entityName, viewName],
    queryFn: () => getDefaultView(entityName, viewName),
    enabled: !!entityName,
  });
}

const parser = new DOMParser();
const serializer = new XMLSerializer();
export const getFetchXml = async (searchText: string, entityLogicalName: string, lookupView?: string, metadata?: IMetadata) => {
  const fetchXmlTemplate = Handlebars.compile(metadata?.associatedView.fetchxml ?? "");
  const fetchXmlMaybeDynamic = metadata?.associatedView.fetchxml ?? "";
  let shouldDefaultSearch = false;
  if (metadata?.associatedView.querytype === 64) {
    shouldDefaultSearch = true;
  } else {
    if (
      !fetchXmlMaybeDynamic.includes("{{PolyLookupSearch}}") &&
      !fetchXmlMaybeDynamic.includes("{{ PolyLookupSearch}}") &&
      !fetchXmlMaybeDynamic.includes("{{PolyLookupSearch }}") &&
      !fetchXmlMaybeDynamic.includes("{{ PolyLookupSearch }}")
    ) {
      shouldDefaultSearch = true;
    }

    const currentRecord = getCurrentRecord();

    return fetchXmlTemplate({
      ...currentRecord,
      PolyLookupSearch: searchText,
    });
  }

  if (shouldDefaultSearch) {
    // if lookup view is not specified and using default lookup fiew,
    // add filter condition to fetchxml to support search
    const fetchXmlForSearch = getFetchXmlForQuery(fetchXmlMaybeDynamic, searchText, metadata.associatedEntity.Attributes, {
      wildcards: "bothWildcard"
    }, createElementAttributes("filter", undefined,
        createElementAttributes("condition", {
          "attribute": metadata?.associatedEntity.PrimaryNameAttribute ?? "",
          "operator": "like",
          "value": `{0}`
        })).outerHTML
    );
    return fetchXmlForSearch;
  }
};

export async function getMetadata(
  currentTable: string | undefined,
  relationshipName: string | undefined,
  relationship2Name: string | undefined,
  associatedViewName: string | undefined
): Promise<IMetadata> {
  if (typeof currentTable === "undefined" || typeof relationshipName === "undefined")
    return Promise.reject(new Error("Invalid arguments"));

  let relationship1: IManyToManyRelationship | IOneToManyRelationship;
  let relationship2: IOneToManyRelationship | undefined;
  let intersectEntityName = "";
  let associatedEntityName: string | undefined;

  if (typeof relationship2Name === "undefined") {
    relationship1 = await getManytoManyRelationShipDefinition(currentTable, relationshipName);
    intersectEntityName = relationship1.IntersectEntityName;
    associatedEntityName =
      relationship1.Entity1LogicalName === currentTable
        ? relationship1.Entity2LogicalName
        : relationship1.Entity1LogicalName;
  } else {
    relationship1 = await getOnetoManyRelationShipDefinition(currentTable, relationshipName);
    intersectEntityName = relationship1.ReferencingEntity;
    relationship2 = await getManytoOneRelationShipDefinition(intersectEntityName, relationship2Name);
    associatedEntityName = relationship2?.ReferencedEntity;
  }

  try {
    const [currentEntity, intersectEntity, associatedEntity, associatedView] = await Promise.all([
      getEntityDefinition(currentTable),
      getEntityDefinition(intersectEntityName),
      getEntityDefinition(associatedEntityName),
      getDefaultView(associatedEntityName, associatedViewName),
    ]);

    let currentIntesectAttribute: string;
    let associatedIntesectAttribute: string;
    let currentEntityNavigationPropertyName: string | undefined;
    let associatedEntityNavigationPropertyName: string | undefined;
    if (isManyToMany(relationship1)) {
      currentIntesectAttribute =
        relationship1.Entity1LogicalName === currentTable
          ? relationship1.Entity1IntersectAttribute
          : relationship1.Entity2IntersectAttribute;

      associatedIntesectAttribute =
        relationship1.Entity1LogicalName === currentTable
          ? relationship1.Entity2IntersectAttribute
          : relationship1.Entity1IntersectAttribute;
    } else {
      currentIntesectAttribute = relationship1.ReferencingAttribute;
      associatedIntesectAttribute = relationship2?.ReferencingAttribute ?? "";
      currentEntityNavigationPropertyName = relationship1.ReferencingEntityNavigationPropertyName;
      associatedEntityNavigationPropertyName = relationship2?.ReferencingEntityNavigationPropertyName;
    }

    return {
      relationship1,
      relationship2,
      currentEntity,
      intersectEntity,
      associatedEntity,
      associatedView,
      currentIntesectAttribute,
      associatedIntesectAttribute,
      currentEntityNavigationPropertyName,
      associatedEntityNavigationPropertyName,
    };
  } catch (error) {
    console.log(error);
    throw error;
  }
}

export async function getMetadataGrid(
  currentTable: string | undefined,
  referencedTable: string | undefined,
  relationshipName: string | undefined,
  clientUrl: string,
): Promise<IMetadata> {
  if (!currentTable || !referencedTable || !relationshipName)
    return Promise.reject(new Error("Invalid arguments"));

  
  const [relationship1, relationship2, associatedView] = await Promise.all([
    getOnetoManyRelationShipDefinition(currentTable, relationshipName),
    getConnectionRelationShipDefinition(referencedTable),
    getDefaultView(referencedTable)
  ]);
  const intersectEntityName = "connection";
  const associatedEntityName: string | undefined = referencedTable;  
  if (!relationship1) { throw new Error(`Relationship ${relationshipName} not found`); }
  if (!relationship2) { throw new Error(`Relationship for connection to ${referencedTable} not found`); }

  try {
    const [currentEntity, intersectEntity, associatedEntity/*, associatedView*/] = await Promise.all([
      getEntityDefinition(currentTable),
      getEntityDefinition(intersectEntityName),
      getEntityDefinition(associatedEntityName),
      // getDefaultView(associatedEntityName, associatedViewName),
    ]);

    const currentIntesectAttribute = relationship1.ReferencingAttribute;
    const associatedIntesectAttribute = relationship2?.ReferencingAttribute ?? "";
    const currentEntityNavigationPropertyName = relationship1.ReferencingEntityNavigationPropertyName;
    const associatedEntityNavigationPropertyName = relationship2?.ReferencingEntityNavigationPropertyName;

    return {
      relationship1,
      relationship2,
      currentEntity,
      intersectEntity,
      associatedEntity,
      associatedView,
      currentIntesectAttribute,
      associatedIntesectAttribute,
      currentEntityNavigationPropertyName,
      associatedEntityNavigationPropertyName,
      clientUrl,
    };
  } catch (error) {
    console.log(error);
    throw error;
  }
}

export function retrieveMultiple(
  entitySetName: string | undefined,
  columns: Array<string | undefined> | undefined,
  filter?: string,
  top?: number,
  orderby?: string,
  extra?: string
) {
  return typeof entitySetName === "undefined" ||
    typeof columns === "undefined" ||
    columns.some((c) => typeof c === "undefined")
    ? Promise.reject(new Error("Invalid entity set name or columns"))
    : axios
        .get<{ value: ComponentFramework.WebApi.Entity[] }>(`/api/data/v${apiVersion}/${entitySetName}`, {
          params: {
            $select: columns.join(","),
            $filter: filter,
            $top: top,
            $orderby: orderby,
          },
        })
        .then((res) => res.data.value);
}

export function retrieveMultipleFetch(
  entitySetName: string | undefined,
  fetchXml: string | undefined,
  page?: number,
  count?: number,
  pagingCookies?: string
) {
  if (!entitySetName || !fetchXml)
    return Promise.reject(new Error("Invalid entity set name or fetchXml"));

  const doc = new DOMParser().parseFromString(fetchXml, "application/xml");
  doc.documentElement.setAttribute("page", page?.toString() ?? "1");
  doc.documentElement.setAttribute("count", count?.toString() ?? "50");
  if (pagingCookies) {
    doc.documentElement.setAttribute("paging-cookie", pagingCookies);
  }

  const newFetchXml = new XMLSerializer().serializeToString(doc);

  return axios
    .get<{ value: ComponentFramework.WebApi.Entity[] }>(`/api/data/v${apiVersion}/${entitySetName}`, {
      headers: {
        Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
      },
      params: {
        fetchXml: encodeURIComponent(newFetchXml),
      },
    })
    .then((res) => res.data.value);
}

export async function retrieveAssociatedRecords(
  currentRecordId: string | undefined,
  intersectEntity: string | undefined,
  intersectEntitySet: string | undefined,
  intersectPrimaryIdAttribute: string | undefined,
  currentIntersectAttribute: string | undefined,
  associatedIntersectAttribute: string | undefined,
  associatedEntity: string | undefined,
  associatedEntityPrimaryIdAttribute: string | undefined,
  associatedEntityPrimaryNameAttribute: string | undefined,
  associatedEntityIconVectorName: string | null | undefined,
  clientUrl?: string,
) {
  if (
    typeof currentRecordId === "undefined" ||
    typeof intersectEntity === "undefined" ||
    typeof intersectEntitySet === "undefined" ||
    typeof intersectPrimaryIdAttribute === "undefined" ||
    typeof currentIntersectAttribute === "undefined" ||
    typeof associatedIntersectAttribute === "undefined" ||
    typeof associatedEntity === "undefined" ||
    typeof associatedEntityPrimaryIdAttribute === "undefined" ||
    typeof associatedEntityPrimaryNameAttribute === "undefined" || 
    typeof associatedEntityIconVectorName === "undefined"
  ) {
    return Promise.reject(new Error("Invalid arguments"));
  }

  const fetchXml = `<fetch>
    <entity name="${intersectEntity}">
      <attribute name="${intersectPrimaryIdAttribute}" />  
      <filter>
        <condition attribute="${currentIntersectAttribute}" operator="eq" value="${currentRecordId}" />
      </filter>
      <link-entity name="${associatedEntity}" from="${associatedEntityPrimaryIdAttribute}" to="${associatedIntersectAttribute}" alias="aLink">
        <attribute name="${associatedEntityPrimaryIdAttribute}" />  
        <attribute name="${associatedEntityPrimaryNameAttribute}" />
      </link-entity>
    </entity>
  </fetch>`;

  const results = await retrieveMultipleFetch(intersectEntitySet, fetchXml);
  return results.map((r) => {
    const id = r[`aLink.${associatedEntityPrimaryIdAttribute}`];
    const name = r[`aLink.${associatedEntityPrimaryNameAttribute}`];
    return {
      entityReference: {
        id: id,
        name: name,
        etn: associatedEntity,
      },
      entityIconUrl: !!associatedEntityIconVectorName ? `${clientUrl ?? ""}/webresources/${associatedEntityIconVectorName}` : null,
      key: id,
      name: name,
      data: {
        [intersectPrimaryIdAttribute]: r[intersectPrimaryIdAttribute],
        [associatedEntityPrimaryIdAttribute]: id,
        [associatedEntityPrimaryNameAttribute]: name,
      } as ComponentFramework.WebApi.Entity
    } as ILookupItem
  });
}

export function associateRecord(
  entitySetName: string | undefined,
  currentRecordId: string | undefined,
  associatedEntitySet: string | undefined,
  associateRecordId: string | undefined,
  relationshipName: string | undefined,
  clientUrl: string | undefined
) {
  if (
    typeof entitySetName === "undefined" ||
    typeof currentRecordId === "undefined" ||
    typeof associatedEntitySet === "undefined" ||
    typeof associateRecordId === "undefined" ||
    typeof relationshipName === "undefined" ||
    typeof clientUrl === "undefined"
  ) {
    return Promise.reject(new Error("Invalid arguments"));
  }

  return axios.post(
    `/api/data/v${apiVersion}/${entitySetName}(${currentRecordId})/${relationshipName}/$ref`,
    {
      "@odata.id": `${clientUrl}/api/data/v${apiVersion}/${associatedEntitySet}(${associateRecordId
        .replace("{", "")
        .replace("}", "")})`,
    },
    {
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "If-None-Match": null,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    }
  );
}

export function disassociateRecord(
  entitySetName: string | undefined,
  currentRecordId: string | undefined,
  relationshipName: string | undefined,
  associatedRecordId: string | undefined
) {
  if (
    typeof entitySetName === "undefined" ||
    typeof currentRecordId === "undefined" ||
    typeof relationshipName === "undefined" ||
    typeof associatedRecordId === "undefined"
  ) {
    return Promise.reject(new Error("Invalid arguments"));
  }

  return axios.delete(
    `/api/data/v${apiVersion}/${entitySetName}(${currentRecordId})/${relationshipName}(${associatedRecordId})/$ref`,
    {
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "If-Match": null,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    }
  );
}

export function createRecord(entitySetName: string | undefined, record: ComponentFramework.WebApi.Entity) {
  if (typeof entitySetName === "undefined") return Promise.reject(new Error("Invalid entity set name"));

  return axios.post(`/api/data/v${apiVersion}/${entitySetName}`, record, {
    headers: {
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Accept: "application/json",
      "Content-Type": "application/json",
    },
  });
}

export function deleteRecord(entitySetName: string | undefined, recordId: string | undefined) {
  if (typeof entitySetName === "undefined" || typeof recordId === "undefined")
    return Promise.reject(new Error("Invalid arguments"));

  return axios.delete(`/api/data/v${apiVersion}/${entitySetName}(${recordId})`, {
    headers: {
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "If-Match": null,
      Accept: "application/json",
      "Content-Type": "application/json",
    },
  });
}

export function getCurrentRecord(attributes?: string[]): ComponentFramework.WebApi.Entity {
  let entityAttributes = Xrm.Page.data.entity.attributes.get();

  if (attributes?.length) {
    entityAttributes = entityAttributes.filter((a) => attributes.includes(a.getName()));
  }

  return Object.fromEntries(
    entityAttributes.map((attribute) => {
      const attributeType = attribute.getAttributeType();
      if (attributeType === "lookup") {
        const lookupVal = (attribute as Xrm.Attributes.LookupAttribute).getValue()?.at(0);
        if (lookupVal) {
          lookupVal.id = lookupVal.id.replace("{", "").replace("}", "");
        }

        return [attribute.getName(), lookupVal];
      } else if (attributeType === "datetime") {
        return [attribute.getName(), (attribute as Xrm.Attributes.DateAttribute).getValue()?.toISOString()];
      }

      return [attribute.getName(), attribute.getValue()];
    })
  );
}
