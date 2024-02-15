import { IInputs } from "../generated/ManifestTypes";

interface IExtendedUserSettings extends ComponentFramework.UserSettings {
  pagingLimit: number;
}

/**
 * Interface for a Lookup value.
 */
interface LookupValue {
  /**
   * The identifier.
   */
  id: string;

  /**
   * The name
   */
  name?: string | undefined;

  /**
   * Type of the entity.
   */
  entityType: string;
}

interface PageInputEntityRecord {
  pageType: "entityrecord";
  /**
   * Logical name of the entity to display the form for.
   * */
  entityName: string;
  /**
   * ID of the entity record to display the form for. If you don't specify this value, the form will be opened in create mode.
   * */
  entityId?: string | undefined;
  /**
   * Designates a record that will provide default values based on mapped attribute values. The lookup object has the following String properties: entityType, id, and name (optional).
   */
  createFromEntity?: LookupValue | undefined;
  /**
   * A dictionary object that passes extra parameters to the form. Invalid parameters will cause an error.
   */
  data?: { [attributeName: string]: any } | undefined;
  /**
   * ID of the form instance to be displayed.
   */
  formId?: string | undefined;
  /**
   * Indicates whether the form is navigated to from a different entity using cross-entity business process flow.
   */
  isCrossEntityNavigate?: boolean | undefined;
  /**
   * Indicates whether there are any offline sync errors.
   */
  isOfflineSyncError?: boolean | undefined;
  /**
   * ID of the business process to be displayed on the form.
   */
  processId?: string | undefined;
  /**
   * ID of the business process instance to be displayed on the form.
   */
  processInstanceId?: string | undefined;
  /**
   * Define a relationship object to display the related records on the form.
   */
  relationship?: any | undefined;
  /**
   * ID of the selected stage in business process instance.
   */
  selectedStageId?: string | undefined;
}

/**
 * Options for navigating to a page: whether to open inline or in a dialog. If you don't specify this parameter, page is opened inline by default.
 * */
interface NavigationOptions {
  /**
   * Specify 1 to open the page inline; 2 to open the page in a dialog.
   * Entity lists can only be opened inline; web resources can be opened either inline or in a dialog.
   * */
  target: 1 | 2;
  /**
   * The width of dialog. To specify the width in pixels, just type a numeric value. To specify the width in percentage, specify an object of type
   * */
  width?: number | SizeValue | undefined;
  /**
   * The width of dialog. To specify the width in pixels, just type a numeric value. To specify the width in percentage, specify an object of type
   * */
  height?: number | SizeValue | undefined;
  /**
   * Specify 1 to open the dialog in center; 2 to open the dialog on the side. Default is 1 (center).
   * */
  position?: 1 | 2 | undefined;
  /*
   * The dialog title on top of the center or side dialog.
   */
  title?: string | undefined;
}

interface SizeValue {
  /**
   * The numerical value
   * */
  value: number;
  /**
   * The unit of measurement. Specify "%" or "px". Default value is "px"
   * */
  unit: "%" | "px";
}

interface IExtendedNavigation extends ComponentFramework.Navigation {
  /**
   * Navigates to the specified page.
   * @param pageInput Input about the page to navigate to. The object definition changes depending on the type of page to navigate to: entity list or HTML web resource.
   * @param navigationOptions Options for navigating to a page: whether to open inline or in a dialog. If you don't specify this parameter, page is opened inline by default.
   */
  navigateTo(
    pageInput: PageInputEntityRecord,
    navigationOptions?: NavigationOptions
  ): Promise<ComponentFramework.NavigationApi.OpenFormSuccessResponse>;

  _customControlProperties: IUnsupportedCustomControlProperties;
}

interface IUnsupportedCustomControlProperties {
  //Implement. Here is an example object I copied from chrome inspector:
  descriptor: IDescriptor
}

interface IDescriptor {
  Parameters: IDescriptorParameters
}

interface IDescriptorParameters {
  AutoExpand?: string,
  ChartGridMode?: string
  DisableMru?: string
  EnableChartPicker?: string
  EnableJumpBar?: string
  EnableQuickFind?: string
  EnableViewPicker?: string
  HeaderColorCode?: string
  IsUserChart?: string
  IsUserView?: string
  RecordsPerPage?: string
  RelationshipName?: string
  TargetEntityType?: string
  ViewId?: string
  ViewIds?: string
  VisualizationId?: string
}

export interface IExtendedContext extends ComponentFramework.Context<IInputs> {
  page: {
    appId: string;
    entityTypeName: string;
    entityId: string;
    isPageReadOnly: boolean;
    getClientUrl: () => string;
  };
  userSettings: IExtendedUserSettings;
  navigation: IExtendedNavigation;
}

export interface IConnectionEntity {
  statecode?: number,
  statuscode?: number,
  record1objecttypecode?: number,
  createdon?: string,
  _relatedconnectionid_value?: ComponentFramework.EntityReference,
  _ownerid_value?: ComponentFramework.EntityReference,
  name?: string,
  ismaster?: true,
  versionnumber?: number,
  connectionid?: string,
  _transactioncurrencyid_value?: ComponentFramework.EntityReference,
  exchangerate?: number,
  _record2id_value?: ComponentFramework.EntityReference,
  record2objecttypecode?: number,
  _modifiedby_value?: ComponentFramework.EntityReference,
  _modifiedonbehalfby_value?: ComponentFramework.EntityReference,
  modifiedon?: string,
  _createdby_value?: ComponentFramework.EntityReference,
  _createdonbehalfby_value?: ComponentFramework.EntityReference,
  _owningbusinessunit_value?: ComponentFramework.EntityReference,
  _record1id_value?: ComponentFramework.EntityReference,
  _owninguser_value?: ComponentFramework.EntityReference,
  _record2roleid_value?: ComponentFramework.EntityReference,
  overriddencreatedon?: string,
  effectiveend?: string,
  entityimage_url?: string,
  entityimageid?: string,
  _owningteam_value?: ComponentFramework.EntityReference,
  effectivestart?: string,
  entityimage_timestamp?: string,
  //entityimage?: string, ???
  _record1roleid_value?: ComponentFramework.EntityReference,
  importsequencenumber?: number,
  description?: string,
}