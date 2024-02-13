import * as React from "react";
import PolyLookupControl, { PolyLookupProps, RelationshipTypeEnum } from "./components/PolyLookupControl";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IExtendedContext } from "./types/extendedContext";
import { LanguagePack } from "./types/languagePack";
import ReactDOM from "react-dom";

export class PolyLookup implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private theComponent: ComponentFramework.StandardControl<IInputs, IOutputs>;
  private notifyOutputChanged: () => void;
  private output: string | undefined;
  private outputSelectedItems: string | undefined;
  private context: ComponentFramework.Context<IInputs>;
  private languagePack: LanguagePack;
  private container?: HTMLDivElement;

  /**
   * Empty constructor.
   */
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    this.languagePack = {
      AddNewLabel: context.resources.getString("AddNewLabel"),
      ControlIsNotAvailableMessage: context.resources.getString("ControlIsNotAvailableMessage"),
      CreateFormNotSupportedMessage: context.resources.getString("CreateFormNotSupportedMessage"),
      EmptyListDefaultMessage: context.resources.getString("EmptyListDefaultMessage"),
      EmptyListMessage: context.resources.getString("EmptyListMessage"),
      LoadingMessage: context.resources.getString("LoadingMessage"),
      PlaceholderDefault: context.resources.getString("PlaceholderDefault"),
      Placeholder: context.resources.getString("Placeholder"),
      RelationshipNotSupportedMessage: context.resources.getString("RelationshipNotSupportedMessage"),
      SuggestionListHeaderDefaultLabel: context.resources.getString("SuggestionListHeaderDefaultLabel"),
      SuggestionListHeaderLabel: context.resources.getString("SuggestionListHeaderLabel"),
      LoadMoreLabel: context.resources.getString("LoadMoreLabel"),
      NoMoreRecordsMessage: context.resources.getString("NoMoreRecordsMessage"),
      SuggestionListFullMessage: context.resources.getString("SuggestionListFullMessage"),
    };
    this.container = container;
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   * @returns ReactElement root react element for the control
   */
  public updateView(context: IExtendedContext): void {
    if (!this.container) { throw "Container must be present" }
    this.context = context;

    let clientUrl = "";

    try {
      clientUrl = context.page.getClientUrl();
    } catch {}

    const props: PolyLookupProps = {
      currentTable: context.page.entityTypeName,
      currentRecordId: context.page.entityId,
      relationshipName: context.parameters.relationship?.raw ?? "",
      relationship2Name: context.parameters.relationship2?.raw ?? undefined,
      relationshipType: Number.parseInt(context.parameters.relationshipType?.raw) as RelationshipTypeEnum,
      clientUrl: clientUrl,
      lookupView: context.parameters.lookupView?.raw ?? undefined,
      itemLimit: context.parameters.itemLimit?.raw ?? undefined,
      pageSize: context.userSettings?.pagingLimit ?? undefined,
      disabled: context.mode.isControlDisabled,
      formType: typeof Xrm === "undefined" ? undefined : Xrm.Page.ui.getFormType(),
      outputSelectedItems: !!context.parameters.outputField?.attributes?.LogicalName,
      defaultLanguagePack: this.languagePack,
      languagePackPath: context.parameters.languagePackPath?.raw ?? undefined,
      onChange:
        context.parameters.outputSelected?.raw === "1" || context.parameters.outputField?.attributes?.LogicalName
          ? this.onLookupChange
          : undefined,
      onQuickCreate: context.parameters.allowQuickCreate?.raw === "1" ? this.onQuickCreate : undefined,
    };
    ReactDOM.render(
      React.createElement(PolyLookupControl, props),
      this.container
    );
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
      boundField: this.output,
      outputField: this.outputSelectedItems,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }

  public onLookupChange = (selectedItems: ComponentFramework.EntityReference[] | undefined) => {
    if (this.context.parameters.outputSelected.raw === "1") {
      this.output = selectedItems?.map((item) => item.name).join(", ");
    }

    if (this.context.parameters.outputField.attributes) {
      if (!selectedItems?.length) {
        this.outputSelectedItems = "";
      } else {
        this.outputSelectedItems = JSON.stringify(selectedItems);
      }
    }
    this.notifyOutputChanged();
  };

  public onQuickCreate = async (
    entityName: string | undefined,
    primaryAttribute: string | undefined,
    value: string | undefined,
    useQuickCreateForm: boolean | undefined
  ) => {
    if (entityName && primaryAttribute) {
      let result: ComponentFramework.NavigationApi.OpenFormSuccessResponse;
      if (useQuickCreateForm) {
        result = await this.context.navigation.openForm(
          {
            entityName: entityName,
            useQuickCreateForm: true,
          },
          {
            [primaryAttribute]: value ?? "",
          }
        );
      } else {
        result = await this.context.navigation.openForm(
          {
            windowPosition: 1,
            entityName: entityName,
          },
          {
            [primaryAttribute]: value ?? "",
          }
        );
      }

      if (result?.savedEntityReference?.length) {
        return result.savedEntityReference[0].id;
      }
    }

    return undefined;
  };
}
