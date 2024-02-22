import { BasePicker, concatStyleSetsWithProps, IBasePickerProps, IBasePickerStyleProps, IBasePickerStyles, Icon, IIconStyles, ISpinnerStyles, IStyle, IStyleFunctionOrObject, Spinner, styled, TagItemSuggestion } from '@fluentui/react';
import { getBasePickerStyles, IBasePickerSuggestionsProps, IPickerItemProps, ITagItemStyleProps, ITagItemStyles, TagPickerBase } from '@fluentui/react/lib/Pickers';
import React from 'react';
import {
  getAxiosInstance,
  retrieveMultipleFetch,
  useLanguagePack,
} from "../services/DataverseService";
import { IMetadata } from '../types/metadata';

import { useMutation, useQuery } from '@tanstack/react-query';
import { LanguagePack } from '../types/languagePack';
import { sprintf } from 'sprintf-js';
import { SuggestionInfo } from './SuggestionInfo';
import { ILookupItem, ILookupItemEntity, ILookupPossibleItems, isLookupItem, isLookupItemEntity, LookupItem } from './LookupItem';
import { getFetchXmlForQuery } from '../services/QueryParser';

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
} & IBasePickerProps<ILookupPossibleItems>

class LookupBaseInternal extends BasePicker<ILookupPossibleItems, ILookupPropsInternal> {
  public static defaultProps = {
    onRenderItem: (props: IPickerItemProps<ILookupPossibleItems>) => <LookupItem {...props} />,
    onRenderSuggestionsItem: (props: ILookupPossibleItems) => <TagItemSuggestion>{props.name}</TagItemSuggestion>,
  };
  constructor(hasPickerProps: ILookupPropsInternal) {
    super(hasPickerProps);

    //Bugfix - somehow I can only remove the first item, but afterwards builtin fluent items.indexOf(item) line of code fails to find same object (probably reference changed, but dunno why)
    const originalRemoveItem = this.removeItem;
    this.removeItem = (item) => {
      const items = this.state.items as ILookupItem[];
      const index = items.findIndex(x => x.key === item.key);
      originalRemoveItem(index >= 0 ? items[index] : item);
    };
  }
}

const toEntityReference = (entity: ComponentFramework.WebApi.Entity, metadata: IMetadata | undefined) => ({
  id: entity[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
  name: entity[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
  etn: metadata?.associatedEntity.LogicalName ?? "",
});

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

const getTextFromItem = (item: ILookupPossibleItems) => item.name;

const getSuggestionTags = (suggestions?: FilterQueryResponse[]) =>
  suggestions?.flatMap(({ result: entities, metadata }) => 
    entities.map(entity => ({
      /* note this key is referencedEntity id - when data will be refreshed, it will be connectionid */
      key: entity[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ?? "",
      name: entity[metadata?.associatedEntity.PrimaryNameAttribute ?? ""] ?? "",
      data: entity,
      entityReference: toEntityReference(entity, metadata),
      entityIconUrl: metadata?.associatedEntity.IconVectorName ? `${metadata?.clientUrl ?? ""}/webresources/${metadata?.associatedEntity.IconVectorName}` : null,
      metadata: metadata,
    } as ILookupItem))
  ) ?? [];

type FilterQueryResponse = {
    metadata?: IMetadata,
    result: ComponentFramework.WebApi.Entity[]
}

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
  lookupEntities,
  ...props
}: ILookupProps
) => {
  const pickerRef = React.useRef<TagPickerBase>(null);
  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);
  const languagePack = loadedLanguagePack ?? defaultLanguagePack;
  const [showIcon, setShowIcon] = React.useState(false);
  const [lookupEntityName, setLookupEntityName] = React.useState<string | undefined>(lookupEntities.length === 1 ? lookupEntities[0] : undefined);
  const entityMetadata = React.useMemo(() => lookupEntityName && metadata ? metadata[lookupEntityName] : undefined, [lookupEntityName, metadata])

  const entitiesQuery = useQuery({
    queryKey: ["entitiesQuery", ...lookupEntities],
    queryFn: () => lookupEntities.map(entity => {
      const m = metadata?.[entity];
      return {
        key: m?.associatedEntity.LogicalName ?? "",
        name: m?.associatedEntity.DisplayNameLocalized ?? "",
        logicalName: m?.associatedEntity.LogicalName ?? "",
        entityIconUrl: !!m?.associatedEntity.IconVectorName ? `${m?.clientUrl ?? ""}/webresources/${m?.associatedEntity.IconVectorName}` : null
      } as ILookupItemEntity
    }),
    enabled: !!metadata && lookupEntities.length > 0 }
  );
  const showEntityFilterSuggestions = React.useCallback(async (): Promise<ILookupPossibleItems[]> => {
    if (lookupEntities.length === 1 || !entitiesQuery.data) {
      return [];
    }
    return entitiesQuery.data;
  }, [entitiesQuery.data, lookupEntities.length]);

    // filter query
  const filterQuery = useMutation({
    mutationFn: async ({ searchText, pageSizeParam, metadata }: { searchText: string; pageSizeParam?: number, metadata?: Record<string, IMetadata> }) => {
      const entitiesToSearch = lookupEntityName ? [lookupEntityName] : lookupEntities;
      return Promise.all(entitiesToSearch.map(async x => ({
        metadata:  metadata?.[x],
        result: await retrieveMultipleFetch(
          getAxiosInstance(metadata?.[x].clientUrl ?? ""),
          metadata?.[x].associatedEntity.EntitySetName,
          getFetchXmlForQuery(metadata?.[x].associatedViewFetchXml ?? "", searchText, undefined, {wildcards: "suffixWildcard"}),
          1,
          pageSizeParam
        ),
      }) as FilterQueryResponse))
    },
  });
  const filterSuggestions = React.useCallback(
    async (filterText: string): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({ searchText: filterText, pageSizeParam: pageSize, metadata: metadata });
      return getSuggestionTags(results);
    },
    [filterQuery, metadata, pageSize]
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
  }), [entityMetadata?.associatedEntity.DisplayCollectionNameLocalized, languagePack.SuggestionListHeaderLabel, languagePack.SuggestionListHeaderDefaultLabel, languagePack.EmptyListMessage, languagePack.EmptyListDefaultMessage, languagePack.AddNewLabel, languagePack.LoadMoreLabel, languagePack.NoMoreRecordsMessage, languagePack.SuggestionListFullMessage, pageSize]);
  const getPlaceholder = React.useCallback(() => {
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
  }, [entityMetadata?.associatedEntity.DisplayCollectionNameLocalized, formType, isDataLoading, isEmpty, languagePack.ControlIsNotAvailableMessage, languagePack.CreateFormNotSupportedMessage, languagePack.LoadingMessage, languagePack.Placeholder, languagePack.PlaceholderDefault, outputSelectedItems]);

  const showMoreSuggestions = React.useCallback(
    async (filterText: string): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({
        searchText: filterText,
        pageSizeParam: (pageSize ?? 50) * 2 + 1,
        metadata: metadata
      });
      return getSuggestionTags(results);
    },
    [filterQuery, metadata, pageSize]
  );

  const onClickEntityFilter = React.useCallback((_: React.MouseEvent<HTMLElement, MouseEvent>, item: ILookupItemEntity) => {
    setLookupEntityName(item.logicalName);
    pickerRef.current?.focus();
  }, []);
  const onRenderSuggestionItem = React.useCallback((item: ILookupPossibleItems) => {
    if (isLookupItem(item)) {
      const infoMap = new Map<string, string>();
      item.metadata.associatedView.layoutjson.Rows?.at(0)?.Cells?.forEach((cell) => {
        let displayValue = item.data[cell.Name + "@OData.Community.Display.V1.FormattedValue"];
        if (!displayValue) {
          displayValue = item.data[cell.Name];
        }
        infoMap.set(cell.Name, displayValue ?? "");
      });
      return <SuggestionInfo infoMap={infoMap} iconUrl={item.entityIconUrl ?? undefined} />
    } else if (isLookupItemEntity(item)) {
      const infoMap = new Map<string, string>();
      infoMap.set("displayName", item.name);
      return <SuggestionInfo infoMap={infoMap} iconUrl={item.entityIconUrl ?? undefined} onClick={x => onClickEntityFilter(x, item)} />
    }
    return <></>
  }, [onClickEntityFilter])

  return (
    <>
      <div style={{ position: 'relative', width: '100%' }}>
        <LookupBaseInternal
          ref={pickerRef}
          selectedItems={selectedItems}
          onResolveSuggestions={filterSuggestions}
          onEmptyResolveSuggestions={showEntityFilterSuggestions}
          onGetMoreResults={showMoreSuggestions}
          onChange={(items) => onChange?.(items?.filter(isLookupItem))}
          onItemSelected={onItemSelected}
          styles={React.useCallback((x: IBasePickerStyleProps) => concatStyleSetsWithProps(x, styles, uciLookupStyle), [styles])}
          theme={theme}
          pickerSuggestionsProps={pickerSuggestionsProps}
          disabled={disabled}
          getTextFromItem={getTextFromItem}
          {...props}
          onRenderItem={React.useCallback((props: IPickerItemProps<ILookupPossibleItems>) => {
            const styles: IStyleFunctionOrObject<ITagItemStyleProps, ITagItemStyles> | undefined = disabled
              ? ({ close: { display: "none" } })
              : undefined;
            //return TagPickerBase.defaultProps.onRenderItem(props);
            const item = props.item as ILookupItem;
            return <LookupItem styles={styles} {...props} item={item} iconUrl={item.entityIconUrl ?? undefined} />;
          }, [disabled])}
          onRenderSuggestionsItem={onRenderSuggestionItem}
          resolveDelay={100}
          inputProps={React.useMemo(() => ({
            placeholder: getPlaceholder(),
            onMouseOver: () => {
              setShowIcon(true);
            },
            onMouseLeave: () => {
              setShowIcon(false);
            }
          }), [getPlaceholder])}
          pickerCalloutProps={React.useMemo(() => ({
            calloutMaxWidth: 500,
          }), [])}
          itemLimit={itemLimit}
        />
        {showIcon && isDataLoading && <Spinner styles={spinnerStyles} /> /* trying to load required data fot lookup */}
        {showIcon && !isDataLoading && <Icon iconName='Search' styles={iconStyles} />}
      </div>
    </>
  );
};


export interface ILookupProps extends Omit<ILookupPropsInternal, 'onResolveSuggestions' | 'onChange'> {
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
  onItemSelected?: (selectedItem?: ILookupPossibleItems | undefined) => ILookupPossibleItems | PromiseLike<ILookupPossibleItems> | null;
  selectedItems?: ILookupPossibleItems[] | undefined;
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