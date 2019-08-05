import * as React from                         'react';
import * as ReactDom from                      'react-dom';
import { Text, Log } from                      '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    PropertyPaneSlider,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneToggle,
    PropertyPaneCheckbox,
    PropertyPaneChoiceGroup,
    IPropertyPaneChoiceGroupOption,
    IPropertyPaneField,
    IPropertyPaneDropdownOption,
    PropertyPaneDropdown
} from                                         '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from   '@microsoft/sp-core-library';
import * as strings from                       'SearchWebPartStrings';
import SearchContainer from                    './components/SearchResultsContainer/SearchResultsContainer';
import ISearchContainerProps from              './components/SearchResultsContainer/ISearchResultsContainerProps';
import { ISearchResultsWebPartProps } from     './ISearchResultsWebPartProps';
import ISearchService from                     '../../services/SearchService/ISearchService';
import MockSearchService from                  '../../services/SearchService/MockSearchService';
import SearchService from                      '../../services/SearchService/SearchService';
import * as moment from                        'moment';
import { Placeholder, IPlaceholderProps } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from                    '@microsoft/sp-core-library';
import LocalizationHelper from                 '../../helpers/LocalizationHelper';
import ResultsLayoutOption from                '../../models/ResultsLayoutOption';
import TemplateService from                    '../../services/TemplateService/TemplateService';
import { PropertyPaneTextDialog } from         '../controls/PropertyPaneTextDialog/PropertyPaneTextDialog';
import { update, isEmpty } from                '@microsoft/sp-lodash-subset';
import MockTemplateService from                '../../services/TemplateService/MockTemplateService';
import BaseTemplateService from                '../../services/TemplateService/BaseTemplateService';

const LOG_SOURCE: string = '[SearchResultsWebPart_{0}]';

export default class SearchResultsWebPart extends BaseClientSideWebPart<ISearchResultsWebPartProps> {

    private _searchService: ISearchService;
    private _templateService: BaseTemplateService;
    private _useResultSource: boolean;
    private _queryKeywords: string;
    private _domElement: HTMLElement;

    /**
     * Used to be able to unregister dynamic data events if the source is updated
     */
    private _lastSourceId: string = undefined;
    private _lastPropertyId: string = undefined;

    /**
     * The template to display at render time
     */
    private _templateContentToDisplay: string;

    constructor() {
        super();

        this._parseRefiners = this._parseRefiners.bind(this);
    }

    /**
     * Determines the group fields for the search settings options inside the property pane
     */
    private _getSearchSettingsFields(): IPropertyPaneField<any>[] {
        // Sets up search settings fields
        const searchSettingsFields: IPropertyPaneField<any>[] = [
            PropertyPaneTextField('queryTemplate', {
                label: strings.QueryTemplateFieldLabel,
                value: this.properties.queryTemplate,
                multiline: true,
                resizable: true,
                placeholder: strings.SearchQueryPlaceHolderText,
                deferredValidationTime: 300,
                disabled: this._useResultSource,
            }),
            PropertyPaneTextField('resultSourceId', {
                label: strings.ResultSourceIdLabel,
                multiline: false,
                onGetErrorMessage: this.validateSourceId.bind(this),
                deferredValidationTime: 300
            }),
            PropertyPaneToggle('enableQueryRules', {
                label: strings.EnableQueryRulesLabel,
                checked: this.properties.enableQueryRules,
            }),
            PropertyPaneTextField('selectedProperties', {
                label: strings.SelectedPropertiesFieldLabel,
                description: strings.SelectedPropertiesFieldDescription,
                multiline: true,
                resizable: true,
                value: this.properties.selectedProperties,
                deferredValidationTime: 300
            }),
            PropertyPaneTextField('refiners', {
                label: strings.RefinersFieldLabel,
                description: strings.RefinersFieldDescription,
                multiline: true,
                resizable: true,
                value: this.properties.refiners,
                deferredValidationTime: 300,
            }),
            PropertyPaneSlider('maxResultsCount', {
                label: strings.MaxResultsCount,
                max: 50,
                min: 1,
                showValue: true,
                step: 1,
                value: 50,
            }),
        ];

        return searchSettingsFields;
    }

    /**
     * Determines the group fields for the search query options inside the property pane
     */
    private _getSearchQueryFields(): IPropertyPaneField<any>[] {
        // Sets up search query fields
        let searchQueryConfigFields: IPropertyPaneField<any>[] = [
            PropertyPaneCheckbox('useSearchBoxQuery', {
                checked: false,
                text: strings.UseSearchBoxQueryLabel,
            })
        ];

        if (this.properties.useSearchBoxQuery) {
          // todo: Fix
        } else {
            searchQueryConfigFields.push(
                PropertyPaneTextField('queryKeywords', {
                    label: strings.SearchQueryKeywordsFieldLabel,
                    description: strings.SearchQueryKeywordsFieldDescription,
                    value: this.properties.useSearchBoxQuery ? '' : this.properties.queryKeywords,
                    multiline: true,
                    resizable: true,
                    placeholder: strings.SearchQueryPlaceHolderText,
                    onGetErrorMessage: this._validateEmptyField.bind(this),
                    deferredValidationTime: 500,
                    disabled: this.properties.useSearchBoxQuery
                })
            );
        }

        return searchQueryConfigFields;
    }

    /**
     * Determines the group fields for styling options inside the property pane
     */
    private _getStylingFields(): IPropertyPaneField<any>[] {

        // Options for the search results layout
        const layoutOptions = [
            {
                iconProps: {
                    officeFabricIconFontName: 'List'
                },
                text: strings.ListLayoutOption,
                key: ResultsLayoutOption.List,
            },
            {
                iconProps: {
                    officeFabricIconFontName: 'Tiles'
                },
                text: strings.TilesLayoutOption,
                key: ResultsLayoutOption.Tiles
            },
            {
                iconProps: {
                    officeFabricIconFontName: 'Code'
                },
                text: strings.CustomLayoutOption,
                key: ResultsLayoutOption.Custom,
            }
        ] as IPropertyPaneChoiceGroupOption[];

        const canEditTemplate = this.properties.externalTemplateUrl && this.properties.selectedLayout === ResultsLayoutOption.Custom ? false : true;

        // Sets up styling fields
        let stylingFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('showBlank', {
                label: strings.ShowBlankLabel,
                checked: this.properties.showBlank,
            }),
            PropertyPaneToggle('showResultsCount', {
                label: strings.ShowResultsCountLabel,
                checked: this.properties.showResultsCount,
            }),
            PropertyPaneToggle('showPaging', {
                label: strings.ShowPagingLabel,
                checked: this.properties.showPaging,
            }),
            PropertyPaneChoiceGroup('selectedLayout', {
                label:'Results layout',
                options: layoutOptions
            }),
            new PropertyPaneTextDialog('inlineTemplateText', {
                dialogTextFieldValue: this._templateContentToDisplay,
                onPropertyChange: this._onCustomPropertyPaneChange.bind(this),
                disabled: !canEditTemplate,
                strings: {
                    cancelButtonText: strings.CancelButtonText,
                    dialogButtonLabel: strings.DialogButtonLabel,
                    dialogButtonText: strings.DialogButtonText,
                    dialogTitle: strings.DialogTitle,
                    saveButtonText: strings.SaveButtonText
                }
            })
        ];

        // Only show the template external URL for 'Custom' option
        if (this.properties.selectedLayout === ResultsLayoutOption.Custom) {
            stylingFields.push(PropertyPaneTextField('externalTemplateUrl', {
                label: strings.TemplateUrlFieldLabel,
                placeholder: strings.TemplateUrlPlaceholder,
                deferredValidationTime: 500,
                onGetErrorMessage: this._onTemplateUrlChange.bind(this)
            }));
        }

        return stylingFields;
    }

    /**
     * Opens the Web Part property pane
     */
    private _setupWebPart() {
        this.context.propertyPane.open();
    }

    /**
     * Checks if a field if empty or not
     * @param value the value to check
     */
    private _validateEmptyField(value: string): string {

        if (!value) {
            return strings.EmptyFieldErrorMessage;
        }

        return '';
    }

    /**
     * Ensures the result source id value is a valid GUID
     * @param value the result source id
     */
    private validateSourceId(value: string): string {
        if(value.length > 0) {
            if (!/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/.test(value)) {
                this._useResultSource = false;
                return strings.InvalidResultSourceIdMessage;
            } else {
                this._useResultSource = true;
            }
        } else {
            this._useResultSource = false;
        }

        return '';
    }

    /**
     * Parses refiners from the property pane value by extracting the refiner managed property and its label in the filter panel.
     * @param rawValue the raw value of the refiner
     */
    private _parseRefiners(rawValue: string) : { [key: string]: string } {

        let refiners = {};

        // Get each configuration
        let refinerKeyValuePair = rawValue.split(',');

        if (refinerKeyValuePair.length > 0) {
            refinerKeyValuePair.map((e) => {

                const refinerValues = e.split(':');
                switch (refinerValues.length) {
                    case 1:
                            // Take the same name as the refiner managed property
                            refiners[refinerValues[0]] = refinerValues[0];
                            break;

                    case 2:
                            // Trim quotes if present
                            refiners[refinerValues[0]] = refinerValues[1].replace(/^'(.*)'$/, '$1');
                            break;
                }
            });
        }

        return refiners;
    }

    /**
     * Get the correct results template content according to the property pane current configuration
     * @returns the template content as a string
     */
    private async _getTemplateContent(): Promise<void> {

        let templateContent = null;

        switch (this.properties.selectedLayout) {
            case ResultsLayoutOption.List:
                templateContent = TemplateService.getListDefaultTemplate();
                break;

            case ResultsLayoutOption.Tiles:
                templateContent = TemplateService.getTilesDefaultTemplate();
                break;

            case ResultsLayoutOption.Custom:

                if (this.properties.externalTemplateUrl) {
                    templateContent = await this._templateService.getFileContent(this.properties.externalTemplateUrl);
                } else {
                    templateContent = this.properties.inlineTemplateText ? this.properties.inlineTemplateText : TemplateService.getBlankDefaultTemplate();
                }

                break;

            default:
            break;
        }

        this._templateContentToDisplay = templateContent;
    }

    /**
     * Custom handler when a custom property pane field is updated
     * @param propertyPath the name of the updated property
     * @param newValue the new value for this property
     */
    private _onCustomPropertyPaneChange(propertyPath: string, newValue: any): void {

        // Stores the new value in web part properties
        update(this.properties, propertyPath, (): any => { return newValue; });

        // Call the default SPFx handler
        this.onPropertyPaneFieldChanged(propertyPath);

        // Refreshes the web part manually because custom fields don't update since sp-webpart-base@1.1.1
        // https://github.com/SharePoint/sp-dev-docs/issues/594
        if (!this.disableReactivePropertyChanges) {
            // The render has to be completed before the property pane to refresh to set up the correct property value
            // so the property pane field will use the correct value for future edit
            this.render();
            this.context.propertyPane.refresh();
        }
    }

    /**
     * Custom handler when the external template file URL
     * @param value the template file URL value
     */
    private async _onTemplateUrlChange(value: string): Promise<String> {

        try {
            // Doesn't raise any error if file is empty (otherwise error message will show on initial load...)
            if(isEmpty(value)) {
                return '';
            }
            // Resolves an error if the file isn't a valid .htm or .html file
            else if(!TemplateService.isValidTemplateFile(value)) {
                return strings.ErrorTemplateExtension;
            }
            // Resolves an error if the file doesn't answer a simple head request
            else {
                await this._templateService.ensureFileResolves(value);
                return '';
            }
        } catch (error) {
            return Text.format(strings.ErrorTemplateResolve, error);
        }
    }

    /**
     * Override the base onInit() implementation to get the persisted properties to initialize data provider.
     */
    protected onInit(): Promise<void> {

        this._domElement = this.domElement;

        // Init the moment JS library locale globally
        const currentLocale = this.context.pageContext.cultureInfo.currentUICultureName;
        moment.locale(currentLocale);

        if (Environment.type === EnvironmentType.Local) {
            this._searchService = new MockSearchService();
            this._templateService = new MockTemplateService();

        } else {

            const lcid = LocalizationHelper.getLocaleId(this.context.pageContext.cultureInfo.currentUICultureName);

            this._searchService = new SearchService(this.context);
            this._templateService = new TemplateService(this.context.spHttpClient);
        }

        // Configure search query settings
        this._useResultSource = false;

        // Set the default search results layout
        this.properties.selectedLayout = this.properties.selectedLayout ? this.properties.selectedLayout: ResultsLayoutOption.List;

        // Make sure the data source will be plugged in correctly when loaded on the page
        // Depending of the component loading order, some sources may be unavailable at this time so that's why we use an event listener

        return super.onInit();
    }

    protected get disableReactivePropertyChanges(): boolean {
        // Set this to true if you don't want the reactive behavior.
        return false;
    }

    protected get isRenderAsync(): boolean {
        return true;
    }

    protected renderCompleted(): void {
        super.renderCompleted();

        let renderElement = null;

        const searchContainer: React.ReactElement<ISearchContainerProps> = React.createElement(
            SearchContainer,
            {
                searchDataProvider: this._searchService,
                queryKeywords: this._queryKeywords,
                maxResultsCount: this.properties.maxResultsCount,
                resultSourceId: this.properties.resultSourceId,
                enableQueryRules: this.properties.enableQueryRules,
                selectedProperties: this.properties.selectedProperties ? this.properties.selectedProperties.replace(/\s|,+$/g, '').split(',') : [],
                refiners: this._parseRefiners(this.properties.refiners),
                showPaging: this.properties.showPaging,
                showResultsCount: this.properties.showResultsCount,
                showBlank: this.properties.showBlank,
                displayMode: this.displayMode,
                templateService: this._templateService,
                templateContent: this._templateContentToDisplay
            } as ISearchContainerProps
        );

        const placeholder: React.ReactElement<IPlaceholderProps> = React.createElement(
            Placeholder,
            {
                iconName: strings.PlaceHolderEditLabel,
                iconText: strings.PlaceHolderIconText,
                description: strings.PlaceHolderDescription,
                buttonLabel: strings.PlaceHolderConfigureBtnLabel,
                onConfigure: this._setupWebPart.bind(this)
            }
        );

        if ((this.properties.queryKeywords && !this.properties.useSearchBoxQuery) ||
            (this.properties.useSearchBoxQuery)) {
            renderElement = searchContainer;
        } else {
            if (this.displayMode === DisplayMode.Edit) {
                renderElement = placeholder;
            } else {
                renderElement = React.createElement('div', null);
            }
        }

        ReactDom.render(renderElement, this._domElement);
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

        return {
            pages: [
                {
                    groups: [
                        {
                            groupName: strings.SearchQuerySettingsGroupName,
                            groupFields: this._getSearchQueryFields()
                        },
                        {
                            groupName: strings.SearchSettingsGroupName,
                            groupFields: this._getSearchSettingsFields()
                        },
                    ],
                    displayGroupsAsAccordion: true
                },
                {
                    groups: [
                        {
                            groupName: strings.StylingSettingsGroupName,
                            groupFields: this._getStylingFields()
                        }
                    ],
                    displayGroupsAsAccordion: true
                }
            ]
        };
    }

    public async onPropertyPaneFieldChanged(changedProperty: string) {

        // Detect if the layout has been changed to custom...
        if (changedProperty === 'inlineTemplateText') {

            // Automatically switch the option to 'Custom' if a default template has been edited
            // (meaning the user started from a the list or tiles template)
            if (this.properties.inlineTemplateText && this.properties.selectedLayout !== ResultsLayoutOption.Custom) {
                this.properties.selectedLayout = ResultsLayoutOption.Custom;

                // Reset also the template URL
                this.properties.externalTemplateUrl = '';
            }
        }

    }

    public async render(): Promise<void> {

        // Configure the provider before the query according to our needs
        this._searchService.resultsCount = this.properties.maxResultsCount;
        this._searchService.queryTemplate = this.properties.queryTemplate;
        this._searchService.resultSourceId = this.properties.resultSourceId;
        this._searchService.enableQueryRules = this.properties.enableQueryRules;

        this._queryKeywords =  this.properties.queryKeywords;

        // If a source is selected, use the value from here
        if (this.properties.useSearchBoxQuery) {
           // todo: Fix
        }

        // Determine the template content to display
        // In the case of an external template is selected, the render is done asynchronously waiting for the content to be fetched
        await this._getTemplateContent();

        this.renderCompleted();
    }
}
