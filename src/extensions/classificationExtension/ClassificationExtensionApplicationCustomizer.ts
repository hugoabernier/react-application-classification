import * as React from "react";
import * as ReactDom from "react-dom";
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderName,
  PlaceholderContent
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "ClassificationExtensionApplicationCustomizerStrings";
import ClassificationHeader from "../../../lib/extensions/classificationExtension/components/ClassificationHeader";
import { IClassificationHeaderProps } from "./components/ClassificationHeader.types";

const LOG_SOURCE: string = "ClassificationExtensionApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IClassificationExtensionApplicationCustomizerProperties {
  ClassificationPropertyBag: string;
  DefaultClassification: string;
  DefaultHandlingUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ClassificationExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IClassificationExtensionApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    console.log("onInit", this.properties);
    if (!this.properties.ClassificationPropertyBag) {
      console.log("Missing required configuration parameters");
      const e: Error = new Error("Missing required configuration parameters");
      console.log(LOG_SOURCE, e);
      return Promise.reject(e);
    }

    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {

    console.log("renderPlaceHolders getting placeholdercontnet");
    const header: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top,
      { onDispose: this._onDispose }
    );

    console.log("renderPlaceHolders getting topPlaceholder");
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose });

      // the extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.log("The header placeholder was not found.");
        return;
      }

      console.log("renderPlaceHolders getting const elem");
      const elem: React.ReactElement<IClassificationHeaderProps> = React.createElement(ClassificationHeader, {
        context: this.context,
        ClassificationPropertyBag: this.properties.ClassificationPropertyBag,
        DefaultClassification: this.properties.DefaultClassification,
        DefaultHandlingUrl: this.properties.DefaultHandlingUrl
      });
      console.log("renderPlaceHolders render");
      ReactDom.render(elem, this._topPlaceholder.domElement);
    }
  }

  private _onDispose(): void {
    // empty
  }
}