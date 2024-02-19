import {
  IBasePickerStyles,
} from "@fluentui/react";
import { QueryClient, QueryClientProvider, useMutation } from "@tanstack/react-query";
import React, { useCallback, useEffect } from "react";
import {
  createRecord,
  deleteRecord,
  useMetadataGrid,
  useSelectedItemsGrid,
} from "../services/DataverseService";
import { LanguagePack } from "../types/languagePack";
import { IMetadata } from "../types/metadata";
import { ILookupItem } from "./LookupItem";
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
  lookupView?: string;
  pageSize?: number;
  disabled?: boolean;
  formType?: XrmEnum.FormType;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  onChange?: (selectedItems: ComponentFramework.EntityReference[] | undefined) => void;
  dataset: ComponentFramework.PropertyTypes.DataSet;
}



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
  const [selectedItemsCreate, setSelectedItemsCreate] = React.useState<ComponentFramework.WebApi.Entity[]>([]);
  const shouldDisable = () => formType !== XrmEnum.FormType.Update;
  const lookupEntitiesArray = React.useMemo(() => lookupEntities?.split(",").map(x => x.trim()) ?? [], [lookupEntities]);
  useEffect(() => {
    console.log("lookupEntitiesArray", lookupEntitiesArray);
  }, [lookupEntitiesArray]);
  if (lookupEntitiesArray.length === 0) {
    //Valid case when there are initially no connections. Don't throw, rather lets find a way to add new entries.
    throw new Error("lookupEntities not set");
  };
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
    mutationFn: (item: ILookupItem) => createRecord(item.metadata.intersectEntity.EntitySetName, {
      [`${item.metadata.currentEntityNavigationPropertyName}@odata.bind`]: `/${item.metadata.currentEntity.EntitySetName}(${currentRecordId})`,
      [`${item.metadata.associatedEntityNavigationPropertyName}@odata.bind`]: `/${item.metadata.associatedEntity.EntitySetName}(${item.entityReference.id})`,
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
    mutationFn: (item: ILookupItem) => deleteRecord(item.metadata?.intersectEntity.EntitySetName, item.connectionReference?.id),
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
    (item?: ILookupItem): ILookupItem | null => {
      return item && !selectedItems?.some(() => item.key === item.key) ? item : null
    },
    [selectedItems]
  );

  const isDataLoading = (isLoadingMetadata || isLoadingSelectedItems || dataset.loading) && !shouldDisable();
  const isEmpty = (((selectedItems?.length == 0 && selectedItemsCreate?.length == 0) ?? true) || disabled) ?? true;
  const metadataObject = React.useMemo(() => Object.values(metadata).every(m => m.isSuccess) ? Object.entries(metadata).reduce((acc, [key, value]) => value.isSuccess ? ({ ...acc, [key]: value.data }) : acc, {}) : undefined, [metadata]);

  return (
    <Lookup 
      metadata={metadataObject}
      formType={formType}
      disabled={disabled}
      selectedItems={selectedItems}
      // pickerSuggestionsProps={pickerSuggestionsProps}
      onChange={onPickerChange}
      onItemSelected={onItemSelected}
      isEmpty={isEmpty}
      defaultLanguagePack={defaultLanguagePack}
      isDataLoading={isDataLoading}
      languagePackPath={languagePackPath}
      pageSize={pageSize}
      lookupView={lookupView}
      // getFetchXml={getFetchXml}
      lookupEntities={lookupEntitiesArray}
      styles={React.useCallback(({ isFocused }) => ({
        root: { backgroundColor: "#fff", width: "100%" },
        input: { minWidth: "0", display: disabled ? "none" : "inline-block" },
        text: {
          minWidth: "0",
          borderColor: "transparent",
          borderWidth: 1,
          borderRadius: 1,
          "&:after": {
            backgroundColor: "transparent",
            borderColor: isFocused ? "#666" : "transparent",
            borderWidth: 1,
            borderRadius: 1,
          },
          "&:hover:after": { backgroundColor: disabled ? "rgba(50, 50, 50, 0.1)" : "transparent" },
        },
    } as Partial<IBasePickerStyles>), [])}
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
