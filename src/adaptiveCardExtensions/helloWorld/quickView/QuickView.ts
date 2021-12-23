import {
  ISPFxAdaptiveCard,
  BaseAdaptiveCardView,
  IActionArguments
} from "@microsoft/sp-adaptive-card-extension-base";
import * as strings from "HelloWorldAdaptiveCardExtensionStrings";
import {
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState,
} from "../HelloWorldAdaptiveCardExtension";

export interface IQuickViewData {
  subTitle: string;
  title: string;
}

export class QuickView extends BaseAdaptiveCardView<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require("./template/QuickViewTemplate.json");
  }


  public onAction(action: IActionArguments): void {    
    if (action.type === 'Submit') {      
      const { id, message } = action.data;
      switch (id) {
        case 'button1': console.log("apretaste el boton 1");
        
        case 'button2': console.log("apretaste el boton 2");
        
          this.setState({
            subTitle: message
          });
          break;
      }
    }
  }
}
