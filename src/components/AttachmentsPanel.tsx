/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import MsgReader, { FieldsData } from "@kenjiuno/msgreader";
export default class AttachmentPanel extends BaseDialog {
    private message: MsgReader;
    private messageFields: FieldsData;
    private messageAttachments: FieldsData[];

    public constructor(msg: MsgReader) {
        super();
        this.message = msg;
        this.messageFields = this.message.getFileData();
        this.messageAttachments = this.messageFields.attachments;
    }
    public render(): void {
        const columns: IColumn[] = [
            {
                key: "fileName",
                name: "fileName",
                minWidth: 300,
                onRender:(item?: FieldsData, index?: number, column?: IColumn) =>{ 
                    debugger;
                    return item.fileName
                } 
            },
        {
                key: "contentLength",
                name: "contentLength",
                minWidth: 100,
                onRender:(item?: FieldsData, index?: number, column?: IColumn) =>{ 
                    debugger;
                    return item.contentLength
                } 
            }

        ]
        ReactDOM.render(<Panel
            isLightDismiss={true}
            isOpen={true}
            headerText="Viewing Attchments"
            isBlocking={false}


            type={PanelType.medium}
            onDismissed={() => this.close()}
        >
            There are {this.messageAttachments.length} attachments
            <DetailsList
                items={this.messageAttachments.filter(item=>{
                    debugger;
                    return !item.attachmentHidden 
                })}
                columns={columns}
            ></DetailsList>
        </Panel>, this.domElement);
    }
}