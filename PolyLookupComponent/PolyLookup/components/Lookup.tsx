import { BasePicker, concatStyleSetsWithProps, IBasePickerProps, IBasePickerStyleProps, IBasePickerStyles, Icon, IIconStyles, IInputProps, ISpinnerStyles, IStyle, IStyleFunctionOrObject, Spinner, styled, TagItem, TagItemSuggestion } from '@fluentui/react';
import { getBasePickerStyles, IBasePickerSuggestionsProps, IPickerItemProps, ITag, ITagItemProps, ITagItemStyleProps, ITagItemStyles, TagPickerBase, ValidationState } from '@fluentui/react/lib/Pickers';
import React from 'react';
import {
  retrieveMultipleFetch,
  useLanguagePack,
} from "../services/DataverseService";
import { IMetadata } from 'PolyLookup/types/metadata';

import { useMutation, UseMutationResult } from '@tanstack/react-query';
import { LanguagePack } from 'PolyLookup/types/languagePack';
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
  onQuickCreate,
  associateQuery,
  disabled,
  onChange,
  onItemSelected,
  selectedItems,
  getFetchXml,
  ...props
}: ILookupProps
) => {
  const pickerRef = React.useRef<TagPickerBase>(null);
  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);
  const languagePack = loadedLanguagePack ?? defaultLanguagePack;
  const [showIcon, setShowIcon] = React.useState(false);
  const filterSuggestions = React.useCallback(
    async (filterText: string, selectedTag?: ILookupItem[]): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({ searchText: filterText, pageSizeParam: pageSize });
      return getSuggestionTags(results, metadata);
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const pickerSuggestionsProps: IBasePickerSuggestionsProps = React.useMemo(() => ({
    suggestionsHeaderText: metadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.SuggestionListHeaderLabel, metadata.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.SuggestionListHeaderDefaultLabel,
    noResultsFoundText:  metadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.EmptyListMessage, metadata.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.EmptyListDefaultMessage,
    forceResolveText: languagePack.AddNewLabel,
    showForceResolve: () => onQuickCreate !== undefined,
    resultsFooter: () => <div>{languagePack.NoMoreRecordsMessage}</div>,
    resultsFooterFull: () => <div>{languagePack.SuggestionListFullMessage}</div>,
    resultsMaximumNumber: (pageSize ?? 50) * 2,
    searchForMoreText: languagePack.LoadMoreLabel,
  }), [languagePack, onQuickCreate, pageSize, metadata]);
  const associatedTableSetName = metadata?.associatedEntity.EntitySetName ?? "";
  
  function getSuggestionTags(
    suggestions: ComponentFramework.WebApi.Entity[] | undefined,
    metadata: IMetadata | undefined
  ) {
    return (
      suggestions?.map(
        (i) =>
          ({
            key: i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ?? "",
            name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""] ?? "",
            data: i,
            entityReference: toEntityReference(i, metadata)
          }) as ILookupItem
      ) ?? []
    );
  }

  // filter query
  const filterQuery = useMutation({
    mutationFn: ({ searchText, pageSizeParam }: { searchText: string; pageSizeParam: number | undefined }) =>
      retrieveMultipleFetch(associatedTableSetName, getFetchXml?.(searchText), 1, pageSizeParam),
  });
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

    return metadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.Placeholder, metadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.PlaceholderDefault;
  };


  const onCreateNew = (input: string): ValidationState => {
    if (onQuickCreate && associateQuery) {
      onQuickCreate(
        metadata?.associatedEntity.LogicalName,
        metadata?.associatedEntity.PrimaryNameAttribute,
        input,
        metadata?.associatedEntity.IsQuickCreateEnabled
      )
        .then((result) => {
          if (result) {
            associateQuery.mutate(result);
            // TODO: fix this hack
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-ignore
            pickerRef.current.input.current?._updateValue("");
          }
        })
        .catch((err: any) => {
          console.log(err);
        });
    }
    return ValidationState.invalid;
  };

  const showMoreSuggestions = React.useCallback(
    async (filterText: string, selectedTag?: ILookupItem[]): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({
        searchText: filterText,
        pageSizeParam: (pageSize ?? 50) * 2 + 1,
      });
      return getSuggestionTags(results, metadata);
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const showAllSuggestions = React.useCallback(
    async (selectedTags?: ILookupItem[]): Promise<ILookupItem[]> => {
      const results = await filterQuery.mutateAsync({ searchText: "", pageSizeParam: pageSize });
      return getSuggestionTags(results, metadata);
    },
    [metadata?.associatedEntity.PrimaryIdAttribute]
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
          onRenderSuggestionsItem={React.useCallback((tag: ILookupItem) => {
            const data = tag.data;
            const infoMap = new Map<string, string>();
            metadata?.associatedView?.layoutjson?.Rows?.at(0)?.Cells.forEach((cell) => {
              let displayValue = data[cell.Name + "@OData.Community.Display.V1.FormattedValue"];
              if (!displayValue) {
                displayValue = data[cell.Name];
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
          onValidateInput={onCreateNew}
        />
        {showIcon && isDataLoading && <Spinner styles={spinnerStyles} /> /* trying to load required data fot lookup */}
        {showIcon && !isDataLoading && <Icon iconName='Search' styles={iconStyles} />}
      </div>
    </>
  );
};


export interface ILookupProps extends Omit<ILookupPropsInternal, 'onResolveSuggestions'> {
  metadata?: IMetadata;
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
  onQuickCreate?: (
    entityName: string | undefined,
    primaryAttribute: string | undefined,
    value: string | undefined,
    useQuickCreateForm: boolean | undefined
  ) => Promise<string | undefined>;
  associateQuery?: UseMutationResult<AxiosResponse<any, any>, unknown, string, unknown>;
  onChange?: (items?: ILookupItem[] | undefined) => void
  onItemSelected?: (selectedItem?: ILookupItem | undefined) => ILookupItem | PromiseLike<ILookupItem> | null;
  selectedItems?: ILookupItem[] | undefined;
  getFetchXml?: (searchText: string) => string | undefined;
}
export const Lookup = styled<ILookupProps, IBasePickerStyleProps, IBasePickerStyles>(
  LookupBase,
  getBasePickerStyles,
  undefined,
  {
    scope: 'Lookup'
  }
);