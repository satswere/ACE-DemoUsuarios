import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
import { IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../HelloWorldAdaptiveCardExtension';

import { IActionArguments } from '@microsoft/sp-adaptive-card-extension-base';

export class CardView extends BasePrimaryTextCardView<IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] {
    const buttons: ICardButton[] = [];

    // Hide the Previous button if at Step 1
   /* if (this.state.currentIndex > 0) {
      buttons.push({
        title: 'Previous',
        action: {
          type: 'Submit',
          parameters: {
            id: 'previous',
            op: -1 // Decrement the index
          }
        }
      });
    }*/  ///botoon para regresar al anterior elemento

    buttons.push({
      title: strings.QuickViewButton,
      action: {
        type: 'QuickView',
        parameters: {
          view: QUICK_VIEW_REGISTRY_ID
        }
      }
    });

    // Hide the Next button if at the end
    if (this.state.currentIndex < this.state.items.length - 1) {
      buttons.push({
        title: 'Next',
        action: {
          type: 'Submit',
          parameters: {
            id: 'next',
            op: 1 // Increment the index
          }
        }
      });
    }
  
    return buttons as [ICardButton] | [ICardButton, ICardButton];
  }
  
  public get data(): IPrimaryTextCardParameters {
   console.log("card normal");
    console.log(this.state.items);
      const { correo, nombre } = this.state.items[this.state.currentIndex];
      return {
        description : correo,
        primaryText: nombre
      };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
         view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }
  public onAction(action: IActionArguments): void {
    if (action.type === 'Submit') {
      const { id, op } = action.data;
      switch (id) {
        case 'previous':
        case 'next':
        this.setState({ currentIndex: this.state.currentIndex + op });
        break;
      }
    }
  }
}
