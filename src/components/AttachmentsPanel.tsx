/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { SPFI } from "@pnp/sp";
import MsgReader, { FieldsData } from "@kenjiuno/msgreader";
export interface IAttachmentPanelProps {
    message: MsgReader;
    sp: SPFI;
}
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
export interface IAttachmentPanelState {

    isOpen: boolean;

}
export default class AttachmentPanel extends React.Component<IAttachmentPanelProps, IAttachmentPanelState> {



    // public constructor(msg: MsgReader) {
    //     super({});
    //     this.message = msg;
    //     this.messageFields = this.message.getFileData();
    //     this.messageAttachments = this.messageFields.attachments;
    // }
    public constructor(props: IAttachmentPanelProps) {
        super(props);
        this.state = { isOpen: true };
    }
    public componentDidMount(): void {
        this.setState({ isOpen: true })
    }
    public render(): React.ReactElement<{}> {
        const columns: IColumn[] = [
            {
                key: "fileName",
                name: "fileName",
                minWidth: 300,
                onRender: (item?: FieldsData, index?: number, column?: IColumn) => {
                    //return item.fileName
                    return <Link
                        onClick={async (e) => {
                            const att = this.props.message.getAttachment(item);
                            console.log(att);
                            const folder = await this.props.sp.web.lists.getByTitle("TemporaryEmailAttachments").rootFolder();
                            console.log(folder);
                            this.props.sp.web.getFolderByServerRelativePath(folder.ServerRelativeUrl).files.addUsingPath(
                                folder.ServerRelativeUrl + "/" + att.fileName, att.content, { Overwrite: true });


                        }}
                    // onClick={async (e) => {
                    //     debugger;
                    //     const att = this.props.message.getAttachment(item);
                    //     console.log(att);

                    //     const folder = await this.props.sp.web.lists.getByTitle("TemporaryEmailAttachments").rootFolder;
                    //     console.log(folder);
                    //     //folder.files.addUsingPath()
                    // }}

                    >{item.fileName}</Link>

                }
            },
            {
                key: "contentLength",
                name: "contentLength",
                minWidth: 100,
                onRender: (item?: FieldsData, index?: number, column?: IColumn) => {
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
            onDismiss={(e) => { this.setState({ isOpen: false }); }}
            type={PanelType.medium}

        >

            <DetailsList
                items={this.props.message.getFileData().attachments.filter(item => {

                    return !item.attachmentHidden
                })}
                columns={columns}
            ></DetailsList>
        </Panel>);
    }
}