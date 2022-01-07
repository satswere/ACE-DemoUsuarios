import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseAdaptiveCardExtension } from "@microsoft/sp-adaptive-card-extension-base";
import { CardView } from "./cardView/CardView";
import { QuickView } from "./quickView/QuickView";
import { HelloWorldPropertyPane } from "./HelloWorldPropertyPane";
import { SPHttpClient } from "@microsoft/sp-http";

import { MediumCardView } from "./cardView/MediumCardView";

import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

const MEDIUM_VIEW_REGISTRY_ID: string = "HelloWorld_MEDIUM_VIEW";

//elementos en la comnfiguracion
export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  listId: string;
  cantidad: number;
}
//personalizados
export interface IHelloWorldAdaptiveCardExtensionState {
  subTitle: string;
  currentIndex: number;
  items: IListItem[];
}

export interface IListItem {
  correo: string;
  nombre: string;
  puestoTrabajo: string;
  telefono: string;
  nombrePila: string;
}

const CARD_VIEW_REGISTRY_ID: string = "HelloWorld_CARD_VIEW";
export const QUICK_VIEW_REGISTRY_ID: string = "HelloWorld_QUICK_VIEW";
//INICIAL
export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      subTitle: "no button clicked",
      currentIndex: 0,
      items: [
        {
          nombre: "Prueba",
          correo: "prueba",
          puestoTrabajo: "prueba",
          telefono: "prueba",
          nombrePila: "prueba",
        },
      ],
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.cardNavigator.register(MEDIUM_VIEW_REGISTRY_ID,() => new MediumCardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID,() => new QuickView());

    return this._fetchData();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return (this.properties.iconProperty || require("./assets/SharePointLogo.svg"));
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloWorld-property-pane'*/
      "./HelloWorldPropertyPane"
    ).then((component) => {
      this._deferredPropertyPane = new component.HelloWorldPropertyPane();
    });
  }
  //que vista se va a cargar
  protected renderCard(): string | undefined {
    return this.cardSize === "Medium" ? MEDIUM_VIEW_REGISTRY_ID : CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string,oldValue: any, newValue: any): void {
    /*if (propertyPath === 'description') {
      this.setState({
        subTitle: newValue
      });
    }*/

    if (propertyPath === "cantidad") {
      this._fetchData();
    }

    if (propertyPath === "listId" && newValue !== oldValue) {
      if (newValue) {
        this._fetchData();
      } else {
        this.setState({ items: [] });
      }
    }
  }

  private _fetchData(): Promise<void> {
    if (this.properties.listId) {
      return this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api("/users")
            .top(this.properties.cantidad)
            .orderby("displayName desc")
            .select("displayName,jobTitle,mail,businessPhones")
            .get()
            .then((res) =>
              res.value.map((item) => {
                return {
                  nombre: item.displayName,
                  correo: item.mail,
                  puestoTrabajo: item.jobTitle,
                  nombrePila: item.givenName,
                  telefono: item.businessPhones[0],
                };
              })
            )
            .then((items) => this.setState({ items }));
        });
    }
    return Promise.resolve();
  }
}
