import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloWorldPropertyPane } from './HelloWorldPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';


import { MediumCardView } from './cardView/MediumCardView';


const MEDIUM_VIEW_REGISTRY_ID: string = 'HelloWorld_MEDIUM_VIEW';


export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  listId: string;
}

export interface IHelloWorldAdaptiveCardExtensionState {
  subTitle: string;
  currentIndex: number;
  items: IListItem[];
}

export interface IListItem {
  title: string;
  description: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloWorld_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloWorld_QUICK_VIEW';

export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      subTitle: "no button clicked",
      currentIndex: 0,
      items: []
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.cardNavigator.register(MEDIUM_VIEW_REGISTRY_ID, () => new MediumCardView());

    return this._fetchData();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloWorld-property-pane'*/
      './HelloWorldPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HelloWorldPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return this.cardSize === 'Medium' ? MEDIUM_VIEW_REGISTRY_ID : CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    /*if (propertyPath === 'description') {
      this.setState({
        subTitle: newValue
      });
    }*/
    if (propertyPath === 'listId' && newValue !== oldValue) {
      if (newValue) {
        this._fetchData();
      } else {
        this.setState({ items: [] });
      }
    }
  }

  private _fetchData(): Promise<void> {
    if (this.properties.listId) {
      return this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/GetById(id='${this.properties.listId}')/items`,
        SPHttpClient.configurations.v1
      )
        .then((response) => response.json())
        .then((jsonResponse) => jsonResponse.value.map(
          (item) => { return { title: item.Title, description: item.Description }; })
          )
        .then((items) => this.setState({ items }));
    }
  
    return Promise.resolve();
  }


}
