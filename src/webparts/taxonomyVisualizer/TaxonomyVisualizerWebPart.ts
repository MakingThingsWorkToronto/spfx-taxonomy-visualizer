import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TaxonomyVisualizerWebPartStrings';
import TopicsExpertise from './components/TaxonomyVisualizer';
import { ITaxonomyVisualizerProps } from './components/ITaxonomyVisualizerProps';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
import { TaxonomyService } from '../../services/TaxonomyService';
import LocalizationHelper from '../../helper/LocalizationHelper';
import { IColumnBreakpoints } from '../../models/IColumnBreakpoints';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface ITaxonomyVisualizerWebPartProps {
  title:string;
  termSetId: string;
  levels:number;
  breakpoints:IColumnBreakpoints[];
  linkTemplate:string;
}

export default class TaxonomyVisualizerWebPart extends BaseClientSideWebPart <ITaxonomyVisualizerWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private _taxonomyService: TaxonomyService;

  public async onInit():Promise<void>{
    
    await super.onInit();
    
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();
    
    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    this._taxonomyService = new TaxonomyService(this.context, this.properties.termSetId);

  }

  public render(): void {
    const element: React.ReactElement<ITaxonomyVisualizerProps> = React.createElement(
      TopicsExpertise,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        theme: this._themeVariant,
        service: this._taxonomyService,
        termSetId: this.properties.termSetId,
        linkTemplate: this.properties.linkTemplate,
        levels: this.properties.levels,
        breakpoints: this.properties.breakpoints,
        lcid: this.getCulture(),
        domElement: this.domElement
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private getCulture():number {
    return LocalizationHelper.getLocaleId(this.context.pageContext.cultureInfo.currentUICultureName);
  }

  /**
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if(propertyPath === "termSetId") {
      this._taxonomyService = new TaxonomyService(this.context, newValue);
    }

    if((propertyPath === "breakpoints" || propertyPath === "termSetId" || propertyPath === "linkTemplate") && newValue) {
      this.render();
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('termSetId', {
                  label: strings.TermSetIdLabel
                }),
                PropertyPaneTextField('linkTemplate', {
                  label: strings.LinkTemplateLabel
                }),
                PropertyFieldCollectionData("breakpoints", {
                  key: "breakpoints",
                  label: strings.EditBreakpointsLabel,
                  enableSorting: true,
                  panelHeader: strings.EditBreakpointsLabel,
                  manageBtnLabel: strings.EditBreakpointsLabel,
                  value: this.properties.breakpoints,
                  fields: [
                    {
                      id: "columns",
                      title: strings.ColumnsLabel,
                      type: CustomCollectionFieldType.number,
                      required: true
                    },
                    {
                      id: "minPixels",
                      title: strings.MaxPixelsLabel,
                      type: CustomCollectionFieldType.number,
                      required: true
                    }   
                  ] 
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
