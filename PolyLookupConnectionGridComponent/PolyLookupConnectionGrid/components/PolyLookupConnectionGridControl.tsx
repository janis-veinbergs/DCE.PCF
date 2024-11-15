import {
  concatStyleSetsWithProps,
  IBasePickerStyleProps,
} from "@fluentui/react";
import { QueryClient, QueryClientProvider, useMutation } from "@tanstack/react-query";
import React, { useCallback } from "react";
import {
  createRecord,
  deleteRecord,
  getAxiosInstance,
  useMetadataGrid,
  useSelectedItemsGrid,
} from "../services/DataverseService";
import { LanguagePack } from "../types/languagePack";
import { IMetadata } from "../types/metadata";
import { ILookupItem, ILookupPossibleItems, isLookupItem } from "./LookupItem";
import { Lookup } from "./Lookup";
import { AxiosError } from "axios";

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      refetchOnWindowFocus: false
    }
  }
});

export enum RelationshipTypeEnum {
  ManyToMany,
  Custom,
  Connection,
}

export interface PolyLookupConnectionGridProps {
  currentTable: string;
  currentRecordId: string;
  relationshipName: string;
  clientUrl: string;
  lookupEntities?: string;
  lookupEntitiesRoles?: string;// [{ entity: string, record1roleid: string, record2roleid: string}];
  /** Roles to use when connecting records. We currently don't have a picker for roles when choosing records, so must be preconfigured with specific role. */
  lookupView?: string;
  pageSize?: number;
  disabled?: boolean;
  formType?: XrmEnum.FormType;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  onChange?: (selectedItems: ComponentFramework.EntityReference[] | undefined) => void;
  dataset: ComponentFramework.PropertyTypes.DataSet;
}

export type EntityConfig = { [key: string]: { record1roleid: string, record2roleid: string } };

const toEntityReference = (entity: ComponentFramework.WebApi.Entity, metadata: IMetadata | undefined) => ({
  id: entity[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
  name: entity[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
  etn: metadata?.associatedEntity.LogicalName ?? "",
});

const Body = ({
  currentTable,
  currentRecordId,
  clientUrl,
  relationshipName,
  lookupEntities,
  lookupEntitiesRoles,
  lookupView,
  pageSize,
  disabled,
  formType,
  defaultLanguagePack,
  languagePackPath,
  dataset
}: PolyLookupConnectionGridProps) => {
  if (!dataset.columns.some(x => x.name === "record2id")) {
    throw new Error(`record2id column is mandatory for grid view ${dataset.getViewId()}`);
  };
  const [selectedItemsCreate] = React.useState<ComponentFramework.WebApi.Entity[]>([]);
  const shouldDisable = () => formType !== XrmEnum.FormType.Update;
  const entityConfig = React.useMemo(() => {
    const guidPairs = lookupEntitiesRoles?.split(";");
    const entities = lookupEntities?.split(",").reduce((prev, cur, idx) => {
      const entity = cur.trim() ?? [];
      const role1 = guidPairs?.[idx]?.split(",")?.[0];
      const role2 = guidPairs?.[idx]?.split(",")?.[1];
      if (!role1) { throw new Error(`Role1 not set for entity ${entity}. Specify correct lookupEntitiesRoles. Item count must match lookupEntities.`); }
      if (!role2) { throw new Error(`Role2 not set for entity ${entity}. Specify correct lookupEntitiesRoles. Item count must match lookupEntities.`); }
      prev[entity] = {
        record1roleid: role1,
        record2roleid: role2,
      };
      return prev;
    }, {} as EntityConfig)
    return entities;
  }, [lookupEntities, lookupEntitiesRoles]);
  if (!entityConfig) {
    throw new Error("lookupEntities not set");
  }

  if (Object.keys(entityConfig).length === 0) {
    //Valid case when there are initially no connections. Don't throw, rather lets find a way to add new entries.
    throw new Error("lookupEntities not set");
  };
  const lookupEntitiesArray = React.useMemo(() => Object.keys(entityConfig), [entityConfig]);
  const metadata = useMetadataGrid(
    currentTable,
    lookupEntitiesArray,
    relationshipName,
    clientUrl,
  );
  const isLoadingMetadata = metadata ? Object.values(metadata).some(x => x.isLoading) : false;

  const {
    data: selectedItems,
    isInitialLoading: isLoadingSelectedItems,
  } = useSelectedItemsGrid(currentTable, relationshipName, lookupEntitiesArray, dataset.records, clientUrl);

  // associate query
  const associateQuery = useMutation({
    mutationFn: (item: ILookupItem) => createRecord(getAxiosInstance(clientUrl), item.metadata.intersectEntity.EntitySetName, {
      [`${item.metadata.currentEntityNavigationPropertyName}@odata.bind`]: `/${item.metadata.currentEntity.EntitySetName}(${currentRecordId})`,
      [`${item.metadata.associatedEntityNavigationPropertyName}@odata.bind`]: `/${item.metadata.associatedEntity.EntitySetName}(${item.entityReference.id})`,
      [`record1roleid@odata.bind`]: `/connectionroles(${entityConfig[item.metadata.associatedEntity.LogicalName].record1roleid})`,
      [`record2roleid@odata.bind`]: `/connectionroles(${entityConfig[item.metadata.associatedEntity.LogicalName].record2roleid})`,
    }),
    onSuccess: () => {
      dataset.refresh();
    },
    onError: (data: AxiosError) => {
      console.error((data.response?.data as any)?.error?.message);
    }
  });

  // disassociate query
  const disassociateQuery = useMutation({
    mutationFn: (item: ILookupItem) => deleteRecord(getAxiosInstance(clientUrl), item.metadata?.intersectEntity.EntitySetName, item.connectionReference?.id),
    onSuccess: () => {
      dataset.refresh();
    },
    onError: (data: AxiosError) => {
      console.error((data.response?.data as any)?.error?.message);
    }
  });


  const onPickerChange = useCallback((selectedTags?: ILookupItem[]): void => {
      const added = selectedTags?.filter(t => !selectedItems?.some(i => i.entityReference.id === t.entityReference.id));
      const removed = selectedItems?.filter(i => !selectedTags?.some(t => i.key === t.key));
      added?.forEach(toEntityReference => associateQuery.mutate(toEntityReference));
      removed?.forEach(toEntityReference => disassociateQuery.mutate(toEntityReference));
    },
    [selectedItems, associateQuery, disassociateQuery]
  );

  const onItemSelected = useCallback(
    (item?: ILookupPossibleItems): ILookupItem | null => item && isLookupItem(item) && !selectedItems?.some((i) => i.key === item.key) ? item : null,
    [selectedItems]
  );
  const styles = React.useCallback((props: IBasePickerStyleProps) => concatStyleSetsWithProps(props, {
      root: { backgroundColor: "#fff", width: "100%" },
      input: { minWidth: "0", display: disabled ? "none" : "inline-block" },
      text: {
        minWidth: "0",
        borderColor: "transparent",
        borderWidth: 1,
        borderRadius: 1,
        "&:after": {
          backgroundColor: "transparent",
          borderColor: props.isFocused ? "#666" : "transparent",
          borderWidth: 1,
          borderRadius: 1,
        },
        "&:hover:after": { backgroundColor: disabled ? "rgba(50, 50, 50, 0.1)" : "transparent" },
      },
  }), [disabled]);


  const isDataLoading = (isLoadingMetadata || isLoadingSelectedItems || dataset.loading) && !shouldDisable();
  const isEmpty = React.useMemo(() => (((selectedItems?.length === 0 && selectedItemsCreate?.length === 0) ?? true) || disabled) ?? true, [selectedItems, selectedItemsCreate, disabled]);
  const metadataObject = React.useMemo(() => Object.values(metadata).every(m => m.isSuccess) ? Object.entries(metadata).reduce((acc, [key, value]) => value.isSuccess ? ({ ...acc, [key]: value.data }) : acc, {}) : undefined, [metadata]);
  return (
    <Lookup 
      metadata={metadataObject}
      formType={formType}
      disabled={disabled}
      selectedItems={selectedItems}
      onChange={onPickerChange}
      onItemSelected={onItemSelected}
      isEmpty={isEmpty}
      defaultLanguagePack={defaultLanguagePack}
      isDataLoading={isDataLoading}
      languagePackPath={languagePackPath}
      pageSize={pageSize}
      lookupView={lookupView}
      lookupEntities={lookupEntitiesArray}
      styles={styles}
    />
  );
};

export default function PolyLookupConnectionGridControl(props: PolyLookupConnectionGridProps) {
  return (
    <QueryClientProvider client={queryClient}>
      <Body {...props} />
    </QueryClientProvider>
  );
}
