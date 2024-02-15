import { IIconProps, IStyleFunctionOrObject, TagItem } from '@fluentui/react';
import { IPickerItemProps, ITag, ITagItemProps, ITagItemStyleProps, ITagItemStyles } from '@fluentui/react/lib/Pickers';
import React from 'react';
import { Image, IImageStyleProps, IImageStyles, ImageFit } from '@fluentui/react/lib/Image';
import { IMetadata } from 'PolyLookupConnectionGrid/types/metadata';

export type EntityReference = {
  id: string;
  etn?: string | undefined;
  name: string;
}


export type ILookupItem = ITag & {
  data: ComponentFramework.WebApi.Entity;
  entityReference: EntityReference;
  entityIconUrl: string | null;
  metadata: IMetadata
}

export type ILookupItemProps = IPickerItemProps<ILookupItem>
  & Pick<ITagItemProps, "className" | "enableTagFocusInDisabledPicker" | "title" | "styles" | "theme">
  & {
    imageUrl?: string
  };
  
export const uciLookupItemStyle = (props: ITagItemStyleProps): Partial<ITagItemStyles> => ({
  ...({
    root: {
      fontWeight: 600,
      height: '28px',
      fontFamily: `SegoeUI-Semibold, "Segoe UI Semibold", "Segoe UI Regular", "Segoe UI";`,
      backgroundColor: 'transparent',
      color: props.theme?.semanticColors.link,
      ':hover': {
        fontWeight: 400,
        fontFamily: `"Segoe UI Regular", SegoeUI, "Segoe UI"`,
        backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered,
        color: props.theme?.semanticColors.link,
        cursor: 'pointer',
        '.ms-Button': {
          display: 'block',
        },
      },
      '.ms-Button': {
        display: 'none',
        backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered
      },
    }, text: {
      padding: '4px 4px 4px 0',
      margin: '0',
      display: 'flex',
      alignItems: 'center',
      '&:hover': {
        backgroundColor: 'rgb(216, 216, 216)',
      }
    }
  })
});
const removeButtonIconProps: IIconProps = {...({
  iconName: 'Cancel',
  styles: (props) => ({
    ...({
      root: [{
        fontWeight: 400,
        lineHeight: '28px',
        ':hover': {
          // backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered,
          backgroundColor: 'rgb(216, 216, 216)',
          margin: '0',
          padding: '0 8px 0 8px',
        },
        '&.ms-Button': {
          // display: 'none',
          backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered
        },
        '&.ms-Button:hover': {
          display: 'block',
          // backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered
          backgroundColor: 'rgb(216, 216, 216)'
        },
        '&.ms-Button-icon': {
          height: '100%',
          // backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered
        },
        // '&.ms-Button-icon:hover': {
        //   // backgroundColor: props.theme?.semanticColors.listItemBackgroundHovered
        //   backgroundColor: 'rgb(216, 216, 216)'
        // },
      }]
  } )
  })
})};
const onClickLookupItem = (event: React.MouseEvent<Element>, item: ILookupItem) => {
  Xrm.Navigation.openForm({
    entityId: item.entityReference.id,
    entityName: item.entityReference.etn,
  });
};
const imageStyles: IStyleFunctionOrObject<IImageStyleProps, IImageStyles> = {
  root: {
    width: '16px',
    height: '16px',
    display: 'flex',
    paddingLeft: '7px',
    paddingRight: '7px'
  }
} 


export const LookupItem: React.FunctionComponent<ILookupItemProps> = React.memo(
  ({ item, 
    index, 
    onRemoveItem, 
    selected, 
    disabled, 
    enableTagFocusInDisabledPicker, 
    className,
    removeButtonAriaLabel,
    title,
    styles,
    imageUrl,
    ...props }: ILookupItemProps) => {
  return (
    
    <TagItem
      onRemoveItem={onRemoveItem}
      selected={selected}
      disabled={disabled}
      enableTagFocusInDisabledPicker={enableTagFocusInDisabledPicker}
      item={{ key: item.entityReference.id, name: item.entityReference.name }}
      styles={uciLookupItemStyle}
      index={index}
      className={className}
      removeButtonAriaLabel={removeButtonAriaLabel}
      title={title}
      removeButtonIconProps={removeButtonIconProps}
    >
      { imageUrl && <Image src={imageUrl} styles={imageStyles} imageFit={ImageFit.centerContain} alt="" /> }
      {/* features linking to item - with middle click opening in new window */}
      <div
        onClick={(event) => onClickLookupItem(event, item)}
        onAuxClick={(event) => event.button === 1 /* Auxiliary button pressed, usually the wheel button or the middle button (if present) */ ? onClickLookupItem(event, item) : undefined}
        style={{
          display: 'flex',
        }}>
          {item.entityReference.name}
      </div>
    </TagItem>
  );
}, (prevProps, nextProps) => prevProps.item?.entityReference.id === nextProps.item?.entityReference.id);
LookupItem.displayName = 'TagItemWithLink';


