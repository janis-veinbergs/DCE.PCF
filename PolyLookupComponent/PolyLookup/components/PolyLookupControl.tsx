import {
  IBasePickerStyles,
  IBasePickerSuggestionsProps,
  TagPickerBase,
} from "@fluentui/react";
import { QueryClient, QueryClientProvider, useMutation } from "@tanstack/react-query";
import Handlebars from "handlebars";
import React, { useCallback } from "react";
import { sprintf } from "sprintf-js";
import {
  associateRecord,
  createRecord,
  deleteRecord,
  disassociateRecord,
  getCurrentRecord,
  useLanguagePack,
  useMetadata,
  useSelectedItems,
} from "../services/DataverseService";
import { LanguagePack } from "../types/languagePack";
import { IMetadata } from "../types/metadata";
import { ILookupItem } from "./LookupItem";
import { Lookup } from "./Lookup";

const queryClient = new QueryClient();
const parser = new DOMParser();
const serializer = new XMLSerializer();

export enum RelationshipTypeEnum {
  ManyToMany,
  Custom,
  Connection,
}

export interface PolyLookupProps {
  currentTable: string;
  currentRecordId: string;
  relationshipName: string;
  relationship2Name?: string;
  relationshipType: RelationshipTypeEnum;
  clientUrl: string;
  lookupView?: string;
  itemLimit?: number;
  pageSize?: number;
  disabled?: boolean;
  formType?: XrmEnum.FormType;
  outputSelectedItems?: boolean;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  onChange?: (selectedItems: ComponentFramework.EntityReference[] | undefined) => void;
  onQuickCreate?: (
    entityName: string | undefined,
    primaryAttribute: string | undefined,
    value: string | undefined,
    useQuickCreateForm: boolean | undefined
  ) => Promise<string | undefined>;
}



const toEntityReference = (entity: ComponentFramework.WebApi.Entity, metadata: IMetadata | undefined) => ({
  id: entity[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
  name: entity[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
  etn: metadata?.associatedEntity.LogicalName ?? "",
});

const Body = ({
  currentTable,
  currentRecordId,
  relationshipName,
  relationship2Name,
  relationshipType,
  clientUrl,
  lookupView,
  itemLimit,
  pageSize,
  disabled,
  formType,
  outputSelectedItems,
  defaultLanguagePack,
  languagePackPath,
  onChange,
  onQuickCreate,
}: PolyLookupProps) => {
  
  const [selectedItemsCreate, setSelectedItemsCreate] = React.useState<ComponentFramework.WebApi.Entity[]>([]);

  const pickerRef = React.useRef<TagPickerBase>(null);

  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: languagePack.SuggestionListHeaderDefaultLabel,
    noResultsFoundText: languagePack.EmptyListDefaultMessage,
    forceResolveText: languagePack.AddNewLabel,
    showForceResolve: () => onQuickCreate !== undefined,
    resultsFooter: () => <div>{languagePack.NoMoreRecordsMessage}</div>,
    resultsFooterFull: () => <div>{languagePack.SuggestionListFullMessage}</div>,
    resultsMaximumNumber: (pageSize ?? 50) * 2,
    searchForMoreText: languagePack.LoadMoreLabel,
  };

  const shouldDisable = () => {
    if (formType === XrmEnum.FormType.Create) {
      if (!outputSelectedItems) {
        return true;
      }
    } else if (formType !== XrmEnum.FormType.Update) {
      return true;
    }
    return false;
  };

  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
  } = useMetadata(
    currentTable,
    relationshipName,
    relationshipType === RelationshipTypeEnum.Custom || relationshipType === RelationshipTypeEnum.Connection
      ? relationship2Name
      : undefined,
    lookupView
  );

  if (metadata && isLoadingMetadataSuccess) {
    pickerSuggestionsProps.suggestionsHeaderText = metadata.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.SuggestionListHeaderLabel, metadata.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.SuggestionListHeaderDefaultLabel;

    pickerSuggestionsProps.noResultsFoundText = metadata.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.EmptyListMessage, metadata.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.EmptyListDefaultMessage;
  }

  const associatedFetchXml = metadata?.associatedView.fetchxml;

  const fetchXmlTemplate = Handlebars.compile(associatedFetchXml ?? "");

  // get selected items
  const {
    data: selectedItems,
    isInitialLoading: isLoadingSelectedItems,
    isSuccess: isLoadingSelectedItemsSuccess,
    refetch: selectedItemsRefetch,
  } = useSelectedItems(metadata, currentRecordId, formType);

  if (isLoadingSelectedItemsSuccess && onChange) {
    onChange(
      selectedItems?.map((i) => toEntityReference(i, metadata))
    );
  }

  // filter query
  const getFetchXml = React.useCallback((searchText: string) => {
      const fetchXmlMaybeDynamic = metadata?.associatedView.fetchxml ?? "";
      let shouldDefaultSearch = false;
      if (!lookupView && metadata?.associatedView.querytype === 64) {
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
        const doc = parser.parseFromString(fetchXmlMaybeDynamic, "application/xml");
        const entities = doc.documentElement.getElementsByTagName("entity");
        for (let i = 0; i < entities.length; i++) {
          const entity = entities[i];
          if (entity.getAttribute("name") === metadata?.associatedEntity.LogicalName) {
            const filter = doc.createElement("filter");
            const condition = doc.createElement("condition");
            condition.setAttribute("attribute", metadata?.associatedEntity.PrimaryNameAttribute ?? "");
            condition.setAttribute("operator", "like");
            condition.setAttribute("value", `%${searchText}%`);
            filter.appendChild(condition);
            entity.appendChild(filter);
          }
        }
        return serializer.serializeToString(doc);
      }
    }, [metadata, fetchXmlTemplate]);

  // associate query
  const associateQuery = useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return associateRecord(
          metadata?.currentEntity.EntitySetName,
          currentRecordId,
          metadata?.associatedEntity?.EntitySetName,
          id,
          relationshipName,
          clientUrl
        );
      } else if (
        relationshipType === RelationshipTypeEnum.Custom ||
        relationshipType === RelationshipTypeEnum.Connection
      ) {
        return createRecord(metadata?.intersectEntity.EntitySetName, {
          [`${metadata?.currentEntityNavigationPropertyName}@odata.bind`]: `/${metadata?.currentEntity.EntitySetName}(${currentRecordId})`,
          [`${metadata?.associatedEntityNavigationPropertyName}@odata.bind`]: `/${metadata?.associatedEntity.EntitySetName}(${id})`,
        });
      }
      return Promise.reject(languagePack.RelationshipNotSupportedMessage);
    },
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });

  // disassociate query
  const disassociateQuery = useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return disassociateRecord(metadata?.currentEntity?.EntitySetName, currentRecordId, relationshipName, id);
      } else if (
        relationshipType === RelationshipTypeEnum.Custom ||
        relationshipType === RelationshipTypeEnum.Connection
      ) {
        return deleteRecord(metadata?.intersectEntity.EntitySetName, id);
      }
      return Promise.reject(languagePack.RelationshipNotSupportedMessage);
    },
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });


  const onPickerChange = useCallback(
    (selectedTags?: ILookupItem[]): void => {
      if (formType === XrmEnum.FormType.Create) {
        const removed = selectedItemsCreate?.filter(
          (i) => !selectedTags?.some((t) => t.entityReference.id === i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""])
        );

        const added = selectedTags?.filter((t) => 
            !selectedItemsCreate?.some((i) => t.entityReference.id === i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""])
        ).map((t) => t.data);

        const oldRemoved = selectedItemsCreate?.filter(
          (o) =>
            !removed?.some(
              (r) =>
                r[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ===
                o[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
            )
        );

        const newSelectedItems = [...oldRemoved, ...(added ?? [])];
        setSelectedItemsCreate(newSelectedItems);

        if (onChange) {
          onChange(
            newSelectedItems?.map((i) => toEntityReference(i, metadata))
          );
        }
      } else if (formType === XrmEnum.FormType.Update) {
        const removed = selectedItems?.filter((i) =>
            !selectedTags?.some((t) => t.entityReference.id === i.data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""])
          )
          .map((i) =>
            relationshipType === RelationshipTypeEnum.ManyToMany
              ? i.data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
              : i.data[metadata?.intersectEntity.PrimaryIdAttribute ?? ""]
          );

        const added = selectedTags
          ?.filter((t) => {
            return !selectedItems?.some(
              (i) => i.entityReference.id === t.data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
            );
          })
          .map((t) => t.key);

        added?.forEach((id) => associateQuery.mutate(id as string));
        removed?.forEach((id) => disassociateQuery.mutate(id as string));
      }
    },
    [selectedItems, selectedItemsCreate, metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const onItemSelected = useCallback(
    (item?: ILookupItem): ILookupItem | null => {
      if (!item) return null;

      if (
        formType === XrmEnum.FormType.Create &&
        !selectedItemsCreate?.some((i) => i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] === item.key)
      ) {
        return item;
      } else if (
        formType === XrmEnum.FormType.Update &&
        !selectedItems?.some((i) => i.key === item.key)
      ) {
        return item;
      }
      return null;
    },
    [selectedItems, metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const isDataLoading = (isLoadingMetadata || isLoadingSelectedItems) && !shouldDisable();

  const isEmpty = (((selectedItems?.length == 0 && selectedItemsCreate?.length == 0) ?? true) || disabled) ?? true;

  return (
    <Lookup 
      metadata={metadata}
      formType={formType}
      disabled={disabled}
      itemLimit={itemLimit}
      selectedItems={selectedItems}
      pickerSuggestionsProps={pickerSuggestionsProps}
      onChange={onPickerChange}
      onItemSelected={onItemSelected}
      isEmpty={isEmpty}
      defaultLanguagePack={defaultLanguagePack}
      isDataLoading={isDataLoading}
      associateQuery={associateQuery}
      onQuickCreate={onQuickCreate}
      languagePackPath={languagePackPath}
      outputSelectedItems={outputSelectedItems}
      pageSize={pageSize}
      lookupView={lookupView}
      getFetchXml={getFetchXml}
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

export default function PolyLookupControl(props: PolyLookupProps) {
  return (
    <QueryClientProvider client={queryClient}>
      <Body {...props} />
    </QueryClientProvider>
  );
}
