/* eslint-disable no-debugger */
import { Log } from "@microsoft/sp-core-library";
import MsgReader from "@kenjiuno/msgreader";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";

import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import AttachmentPanel from "../../components/AttachmentsPanel";
import * as React from "react";
import * as ReactDOM from "react-dom";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMessageAttachmentsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "MessageAttachmentsCommandSet";

export default class MessageAttachmentsCommandSet extends BaseListViewCommandSet<IMessageAttachmentsCommandSetProperties> {
  private sp: SPFI;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized MessageAttachmentsCommandSet");
    this.sp = spfi().using(SPFx(this.context));
    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "VIEW_ATTACHMENTS":
        this.viewAttachments(event)
        .catch((e) => {
          debugger;
        });

        break;

      default:
        throw new Error("Unknown command");
    }
  }
  private async viewAttachments(
    event: IListViewCommandSetExecuteEventParameters
  ): Promise<void> {
    const itemId = event.selectedRows[0].getValueByName("ID");
    this.sp.web.lists
      .getById(this.context.pageContext.list.id.toString())
      .items.getById(parseInt(itemId))
      .file()
      .then(async (fileInfo) => {
          const url = fileInfo.ServerRelativeUrl;
        const buffer: ArrayBuffer = await this.sp.web
          .getFileByServerRelativePath(url)
          .getBuffer();
        const messgage = new MsgReader(buffer);
        const div = document.createElement("div");

        const element: React.ReactElement<{}> = React.createElement(
          AttachmentPanel,
          { message: messgage,sp:this.sp,headerText: event.selectedRows[0].getValueByName("FileLeafRef") }
        );
        ReactDOM.render(element, div);
      })
      .catch((e) => {
        debugger;
      });
  }
  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");
    
    const compareOneCommand: Command = this.tryGetCommand("VIEW_ATTACHMENTS");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected and its a .msg
      if (this.context.listView.selectedRows?.length === 1) {
        const filename: string =
          this.context.listView.selectedRows[0].getValueByName("FileLeafRef");
        const filenameparts = filename.split(".");
        compareOneCommand.visible = filenameparts.pop().toLowerCase() === "msg";
      }
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  };
}
