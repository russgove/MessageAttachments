/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import MsgReader, { FieldsData } from "@kenjiuno/msgreader";
export interface IAttachmentPanelProps{
     message: MsgReader;
     isOpen:boolean;
}
export interface IAttachmentPanelState{

    isOpen:boolean;
}
export default class AttachmentPanel extends React.Component<IAttachmentPanelProps,IAttachmentPanelState> {

 

    // public constructor(msg: MsgReader) {
    //     super({});
    //     this.message = msg;
    //     this.messageFields = this.message.getFileData();
    //     this.messageAttachments = this.messageFields.attachments;
    // }
    public constructor(props:IAttachmentPanelProps) {
        super(props);
        this.state={isOpen:true};
    }
    public componentDidMount(): void {
        this.setState({isOpen:true})
    }
    public render(): React.ReactElement<{}> {
        const columns: IColumn[] = [
            {
                key: "fileName",
                name: "fileName",
                minWidth: 300,
                onRender:(item?: FieldsData, index?: number, column?: IColumn) =>{ 
                              return item.fileName
                } 
            },
        {
                key: "contentLength",
                name: "contentLength",
                minWidth: 100,
                onRender:(item?: FieldsData, index?: number, column?: IColumn) =>{ 
                    return item.contentLength
                } 
            }

        ]
        return (<Panel
            isLightDismiss={true}
            isOpen={this.state.isOpen}
            headerText="Viewing Attchments"
            isBlocking={false}
            closeButtonAriaLabel='Close'
onDismiss={(e)=>{this.setState({isOpen:false});}}
            type={PanelType.medium}
            
        >
    
            <DetailsList
                items={this.props.message.getFileData().attachments.filter(item=>{
 
                    return !item.attachmentHidden 
                })}
                columns={columns}
            ></DetailsList>
        </Panel>);
    }
}