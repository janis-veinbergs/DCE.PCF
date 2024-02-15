import { BasePicker, concatStyleSetsWithProps, IBasePickerProps, IBasePickerStyleProps, IBasePickerStyles, Icon, IIconStyles, IInputProps, ISpinnerStyles, IStyle, IStyleFunctionOrObject, Spinner, styled, TagItem, TagItemSuggestion } from '@fluentui/react';
import { getBasePickerStyles, IBasePickerSuggestionsProps, IPickerItemProps, ITag, ITagItemProps, ITagItemStyleProps, ITagItemStyles, TagPickerBase, ValidationState } from '@fluentui/react/lib/Pickers';
import React from 'react';
import {
  retrieveMultipleFetch,
  useLanguagePack,
} from "../services/DataverseService";
import { IMetadata, IViewDefinition } from 'PolyLookupConnectionGrid/types/metadata';

import { useMutation, UseMutationResult } from '@tanstack/react-query';
import { LanguagePack } from 'PolyLookupConnectionGrid/types/languagePack';
import { sprintf } from 'sprintf-js';
import { AxiosResponse } from 'axios';
import { SuggestionInfo } from './SuggestionInfo';
import { ILookupItem, LookupItem } from './LookupItem';

type ILookupPropsInternal = {
  /** Number of results for autocomplete to be returned */
  quickFindCount?: number,
  /** Search against CRM will be issued only when you stop typing search query for this number of milliseconds. */
  onResolveSuggestionsDebounceWait?: number,
  /** Show error message with the control */
  errorMessage?: string | JSX.Element,
  /** Add custom filter in addition to the filter already present in the view to be used for searching. Pass <filter type="and/or"> tag */
  additionalFilter?: string
  /** any additional attributes you want to include within FetchXml. Only single <entity> element within FetchXml supported. Only primary entity attributes, linked entity attributes cannot be added. Use additionalLinkAttributes to add additional <link-entity> with respective attributes */
  additionalAttributes?: string[]
  /** Add additional link-entity. Useful if want to fetch additional attribute from linked entity. to/from entities can be duplicate/alreayd existing under entity element and all specified attributes will be fetched.
  * Example: ['<link-entity name="account" from="accountid" to="deac_accountid" visible="false" link-type="outer" alias="accountnamealias"><attribute name="name" /></link-entity>'] */
  additionalLinkAttributes?: string[]
} & IBasePickerProps<ILookupItem>

class LookupBaseInternal extends BasePicker<ILookupItem, ILookupPropsInternal> {
  public static defaultProps = {
    onRenderItem: (props: IPickerItemProps<ILookupItem>) => <LookupItem {...props} />,
    onRenderSuggestionsItem: (props: ILookupItem) => <TagItemSuggestion>{props.entityReference.name}</TagItemSuggestion>,
  };
}

const toEntityReference = (entity: ComponentFramework.WebApi.Entity, metadata: IMetadata | undefined) => ({
  id: entity[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
  name: entity[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
  etn: metadata?.associatedEntity.LogicalName ?? "",
});

const onClickLookupItem = (event: React.MouseEvent<Element>, item: ILookupItem, options?: Pick<Xrm.Navigation.EntityFormOptions, 'openInNewWindow'>) => {
  Xrm.Navigation.openForm({
    entityId: item.entityReference.id,
    entityName: item.entityReference.etn,
    openInNewWindow: options?.openInNewWindow
  });
};

const uciLookupStyle = (props: IBasePickerStyleProps): Partial<IBasePickerStyles> => ({
  ...(props.disabled ? {
    root: {
      width: '100%',
    },
    text: {
      fontWeight: 600,
      border: 'none',
    },
    itemsWrapper: {
      backgroundColor: 'transparent',
    }
  } : {
    root: {
      width: '100%',
    },
    text: {
      fontWeight: 600,
      backgroundColor: props.theme?.semanticColors.inputBackground,
      borderColor: 'transparent',
      ':after': {
        border: 'none'
      },
    },
  })
});
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const iconStyle: IStyle = {
  position: 'absolute',
  top: 8,
  right: 8,
  pointerEvents: 'none'
};
const spinnerStyles: ISpinnerStyles = { root: iconStyle }
const iconStyles: IIconStyles = { root: iconStyle }

const getTextFromItem = (item: ILookupItem) => item.entityReference.name;

const LookupBase: React.FunctionComponent<ILookupProps> = ({
  styles,
  theme,
  metadata,
  formType,
  pageSize,
  outputSelectedItems,
  defaultLanguagePack,
  languagePackPath,
  isDataLoading,
  isEmpty,
  itemLimit,
  disabled,
  onChange,
  onItemSelected,
  selectedItems,
  getFetchXml,
  lookupEntities,
  ...props
}: ILookupProps
) => {
  const pickerRef = React.useRef<TagPickerBase>(null);
  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);
  const languagePack = loadedLanguagePack ?? defaultLanguagePack;
  const [showIcon, setShowIcon] = React.useState(false);
  //TODO: Allow picking which entity to search for
  const [lookupEntityName, setLookupEntityName] = React.useState(lookupEntities[0]);
  const [entityMetadata, setEntityMetadata] = React.useState(metadata?.[lookupEntityName]);

    // filter query
  const filterQuery = useMutation({
    mutationFn: ({ searchText, pageSizeParam }: { searchText: string; pageSizeParam: number | undefined }) =>
      retrieveMultipleFetch(associatedTableSetName, getFetchXml?.(searchText, lookupEntityName), 1, pageSizeParam),
  });
  const filterSuggestions = React.useCallback(
    async (filterText: string, selectedTag?: ILookupItem[]): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({ searchText: filterText, pageSizeParam: pageSize });
      return getSuggestionTags(results, entityMetadata);
    },
    [entityMetadata?.associatedEntity.EntitySetName, filterQuery]
  );

  const pickerSuggestionsProps: IBasePickerSuggestionsProps = React.useMemo(() => ({
    suggestionsHeaderText: entityMetadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.SuggestionListHeaderLabel, entityMetadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.SuggestionListHeaderDefaultLabel,
    noResultsFoundText:  entityMetadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.EmptyListMessage, entityMetadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.EmptyListDefaultMessage,
    forceResolveText: languagePack.AddNewLabel,
    resultsFooter: () => <div>{languagePack.NoMoreRecordsMessage}</div>,
    resultsFooterFull: () => <div>{languagePack.SuggestionListFullMessage}</div>,
    resultsMaximumNumber: (pageSize ?? 50) * 2,
    searchForMoreText: languagePack.LoadMoreLabel,
  }), [languagePack, pageSize, metadata]);
  const associatedTableSetName = entityMetadata?.associatedEntity.EntitySetName ?? "";
  
  function getSuggestionTags(
    suggestions: ComponentFramework.WebApi.Entity[] | undefined,
    metadata: IMetadata | undefined
  ) {
    return (
      suggestions?.map(
        (i) =>
          ({
            key: i[entityMetadata?.associatedEntity.PrimaryIdAttribute ?? ""] ?? "",
            name: i[entityMetadata?.associatedEntity.PrimaryNameAttribute ?? ""] ?? "",
            data: i,
            entityReference: toEntityReference(i, metadata),
            metadata: metadata,
          }) as ILookupItem
      ) ?? []
    );
  }

  const getPlaceholder = () => {
    if (formType === XrmEnum.FormType.Create) {
      if (!outputSelectedItems) {
        return languagePack.CreateFormNotSupportedMessage;
      }
    } else if (formType !== XrmEnum.FormType.Update) {
      return languagePack.ControlIsNotAvailableMessage;
    }

    if (isDataLoading) {
      return languagePack.LoadingMessage;
    }

    if (isEmpty) {
      return "---";
    }

    return entityMetadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.Placeholder, entityMetadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.PlaceholderDefault;
  };

  const showMoreSuggestions = React.useCallback(
    async (filterText: string, selectedTag?: ILookupItem[]): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({
        searchText: filterText,
        pageSizeParam: (pageSize ?? 50) * 2 + 1,
      });
      return getSuggestionTags(results, entityMetadata);
    },
    [entityMetadata?.associatedEntity.EntitySetName, filterQuery]
  );

  const showAllSuggestions = React.useCallback(
    async (selectedTags?: ILookupItem[]): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({ searchText: "", pageSizeParam: pageSize });
      return getSuggestionTags(results, entityMetadata);
    },
    [entityMetadata?.associatedEntity.PrimaryIdAttribute, filterQuery]
  );

  return (
    <>
      <div style={{ position: 'relative', width: '100%' }}>
        <LookupBaseInternal
          ref={pickerRef}
          selectedItems={selectedItems}
          onResolveSuggestions={filterSuggestions}
          onEmptyResolveSuggestions={showAllSuggestions}
          onGetMoreResults={showMoreSuggestions}
          onChange={onChange}
          onItemSelected={onItemSelected}
          styles={React.useCallback(x => concatStyleSetsWithProps(x, styles, uciLookupStyle), [styles])}
          theme={theme}
          pickerSuggestionsProps={pickerSuggestionsProps}
          disabled={disabled}
          getTextFromItem={getTextFromItem}
          {...props}
          onRenderItem={(props) => {
            const styles: IStyleFunctionOrObject<ITagItemStyleProps, ITagItemStyles> | undefined = disabled
              ? ({ close: { display: "none" } })
              : undefined;
            //return TagPickerBase.defaultProps.onRenderItem(props);
            const item = props.item as ILookupItem;
            return <LookupItem styles={styles} {...props} item={item} imageUrl={item.entityIconUrl ?? undefined} />;
          }}
          onRenderSuggestionsItem={React.useCallback((item: ILookupItem) => {
            const infoMap = new Map<string, string>();
            //useDefaultView(entityLogicalName, lookupView).data
            item.metadata.associatedView.layoutjson.Rows?.at(0)?.Cells?.forEach((cell) => {
              let displayValue = item.data[cell.Name + "@Oitem..Community.Display.V1.FormattedValue"];
              if (!displayValue) {
                displayValue = item.data[cell.Name];
              }
              infoMap.set(cell.Name, displayValue ?? "");
            });
            return <SuggestionInfo infoMap={infoMap}></SuggestionInfo>;
          }, [metadata])}
          resolveDelay={100}
          inputProps={{
            placeholder: getPlaceholder(),
            onMouseOver: () => {
              setShowIcon(true);
            },
            onMouseLeave: () => {
              setShowIcon(false);
            }
          }}
          pickerCalloutProps={{
            calloutMaxWidth: 500,
          }}
          itemLimit={itemLimit}
        />
        {showIcon && isDataLoading && <Spinner styles={spinnerStyles} /> /* trying to load required data fot lookup */}
        {showIcon && !isDataLoading && <Icon iconName='Search' styles={iconStyles} />}
      </div>
    </>
  );
};


export interface ILookupProps extends Omit<ILookupPropsInternal, 'onResolveSuggestions'> {
  metadata?: Record<string, IMetadata>;
  formType?: XrmEnum.FormType;
  lookupView?: string;
  pageSize?: number;
  outputSelectedItems?: boolean;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  isDataLoading: boolean;
  isEmpty: boolean;
  itemLimit?: number;
  disabled?: boolean;
  onChange?: (items?: ILookupItem[] | undefined) => void
  onItemSelected?: (selectedItem?: ILookupItem | undefined) => ILookupItem | PromiseLike<ILookupItem> | null;
  selectedItems?: ILookupItem[] | undefined;
  getFetchXml?: (searchText: string, entityLogicalName: string) => string | undefined;
  lookupEntities: string[];
}
export const Lookup = styled<ILookupProps, IBasePickerStyleProps, IBasePickerStyles>(
  LookupBase,
  getBasePickerStyles,
  undefined,
  {
    scope: 'Lookup'
  }
);