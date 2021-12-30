import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
import { IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState } from '../HelloWorldAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
  description: string;
  profileImage: string;
  correo: string;
  puestoTrabajo: string;
  telefono: string;
  nombrePila: string;
  nombre : string;
}

export class QuickView extends BaseAdaptiveCardView<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    const { correo, nombre,nombrePila,puestoTrabajo,telefono } = this.state.items[this.state.currentIndex];
    return {
      subTitle: "correo",
      correo,
      puestoTrabajo,
      telefono,
      profileImage: "https://www.jeancoutu.com/globalassets/revamp/photo/conseils-photo/20160302-01-reseaux-sociaux-profil/photo-profil_301783868.jpg",
      title: "titulo",
      nombrePila,
      nombre,
      description: this.properties.description
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}