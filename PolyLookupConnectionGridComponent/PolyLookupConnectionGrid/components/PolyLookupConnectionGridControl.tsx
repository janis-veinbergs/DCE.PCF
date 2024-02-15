import {
  IBasePickerStyles,
  IBasePickerSuggestionsProps,
} from "@fluentui/react";
import { QueryClient, QueryClientProvider, useMutation } from "@tanstack/react-query";
import React, { useCallback } from "react";
import {
  createRecord,
  deleteRecord,
  useLanguagePack,
  useMetadataGrid,
  useSelectedItemsGrid,
} from "../services/DataverseService";
import { LanguagePack } from "../types/languagePack";
import { IMetadata } from "../types/metadata";
import { ILookupItem } from "./LookupItem";
import { Lookup } from "./Lookup";
import { useForceUpdate } from "@fluentui/react-hooks";
import { AxiosError } from "axios";

const queryClient = new QueryClient();

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
  onChange,
  dataset
}: PolyLookupConnectionGridProps) => {
  if (!dataset.columns.some(x => x.name === "record2id")) {
    throw new Error(`record2id column is mandatory for grid view ${dataset.getViewId()}`);
  };
  const forceUpdate = useForceUpdate();
  const [selectedItemsCreate, setSelectedItemsCreate] = React.useState<ComponentFramework.WebApi.Entity[]>([]);

  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: languagePack.SuggestionListHeaderDefaultLabel,
    noResultsFoundText: languagePack.EmptyListDefaultMessage,
    forceResolveText: languagePack.AddNewLabel,
    resultsFooter: () => <div>{languagePack.NoMoreRecordsMessage}</div>,
    resultsFooterFull: () => <div>{languagePack.SuggestionListFullMessage}</div>,
    resultsMaximumNumber: (pageSize ?? 50) * 2,
    searchForMoreText: languagePack.LoadMoreLabel,
  };

  const shouldDisable = () => formType !== XrmEnum.FormType.Update;

  // const unique = (value: any, index: number, array: any[]) => array.indexOf(value) === index;
  // const entitiesReferenced = Object.entries(records)
  //   .map(([key, value]) => (value as IConnectionEntity)._record2id_value?.etn)
  //   .filter(unique);
  const lookupEntitiesArray = lookupEntities?.split(",").map(x => x.trim()) ?? [];
  console.log("lookupEntitiesArray", lookupEntitiesArray);
  if (lookupEntitiesArray.length === 0) {
    //Valid case when there are initially no connections. Don't throw, rather lets find a way to add new entries.
    throw new Error("lookupEntities not set");
  };
  const metadata = useMetadataGrid(
    currentTable,
    lookupEntitiesArray,
    relationshipName
  );
  const isLoadingMetadata = metadata ? Object.values(metadata).some(x => x.isLoading) : false;
  const isLoadingMetadataSuccess = metadata ? Object.values(metadata).every(x => x.isSuccess) : false;


    // if (metadata && isLoadingMetadataSuccess) {
  //   pickerSuggestionsProps.suggestionsHeaderText = metadata.associatedEntity.DisplayCollectionNameLocalized
  //     ? sprintf(languagePack.SuggestionListHeaderLabel, metadata.associatedEntity.DisplayCollectionNameLocalized)
  //     : languagePack.SuggestionListHeaderDefaultLabel;

  //   pickerSuggestionsProps.noResultsFoundText = metadata.associatedEntity.DisplayCollectionNameLocalized
  //     ? sprintf(languagePack.EmptyListMessage, metadata.associatedEntity.DisplayCollectionNameLocalized)
  //     : languagePack.EmptyListDefaultMessage;
  // }

  const {
    data: selectedItems,
    isInitialLoading: isLoadingSelectedItems,
  } = useSelectedItemsGrid(currentTable, relationshipName, lookupEntitiesArray, dataset.records);

  // associate query
  const associateQuery = useMutation({
    mutationFn: (item: ILookupItem) => createRecord(item.metadata.intersectEntity.EntitySetName, {
      [`${item.metadata.currentEntityNavigationPropertyName}@odata.bind`]: `/${item.metadata.currentEntity.EntitySetName}(${currentRecordId})`,
      [`${item.metadata.associatedEntityNavigationPropertyName}@odata.bind`]: `/${item.metadata.associatedEntity.EntitySetName}(${item.entityReference.id})`,
    }),
    onSuccess: () => {
      // selectedItemsRefetch();
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
      // selectedItemsRefetch();
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
    (item?: ILookupItem): ILookupItem | null => item && !selectedItems?.some((i) => i.key === item.key) ? item : null,
    [selectedItems]
  );

  const isDataLoading = (isLoadingMetadata || isLoadingSelectedItems || dataset.loading) && !shouldDisable();

  const isEmpty = (((selectedItems?.length == 0 && selectedItemsCreate?.length == 0) ?? true) || disabled) ?? true;

  return (
    <Lookup 
      metadata={React.useMemo(() => Object.entries(metadata).reduce((acc, [key, value]) => ({ ...acc, [key]: value.data }), {}), [metadata])}
      formType={formType}
      disabled={disabled}
      selectedItems={selectedItems}
      pickerSuggestionsProps={pickerSuggestionsProps}
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
