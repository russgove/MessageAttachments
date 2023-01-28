/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { TextField, } from 'office-ui-fabric-react/lib/TextField';
import { Stack, StackItem ,IStackTokens} from 'office-ui-fabric-react/lib/Stack';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { PrimaryButton ,DefaultButton} from 'office-ui-fabric-react/lib/Button';
import { SPFI } from "@pnp/sp";
import MsgReader, { FieldsData } from "@kenjiuno/msgreader";
export interface IAttachmentPanelProps {
    message: MsgReader;
    sp: SPFI;
    headerText: string;
}
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IFileAddResult } from '@pnp/sp/files';
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
        const itemAlignmentsStackTokens: IStackTokens = {
            childrenGap: 5,
            padding: 10,
          };
        console.dir(this.props.message.getFileData());
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
                            const addResult: IFileAddResult = await this.props.sp.web.getFolderByServerRelativePath(folder.ServerRelativeUrl).files.addUsingPath(
                                folder.ServerRelativeUrl + "/" + att.fileName, att.content, { Overwrite: true });
                            console.log(addResult);


                        }}
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

        ];
        debugger;
        const toEmails = this.props.message.getFileData().recipients
            .filter(r => r.recipType === "to")
            .map(r => `${r.name} <${r.smtpAddress}>`)
            .join(" ");

        const ccEmails = this.props.message.getFileData().recipients
            .filter(r => r.recipType === "cc")
            .map(r => `${r.name} <${r.smtpAddress}>`)
            .join(" ");

        const att = this.props.message.getFileData().attachments
            .filter(item => {

                return !item.attachmentHidden
            }).map(att => {
                return (
                    <StackItem>
                        <DefaultButton className='ms-bgColor-themeLight'
                            onClick={async (e) => {
                                debugger;
                                console.log(att);
                                const attchment = this.props.message.getAttachment(att);
                                const folder = await this.props.sp.web.lists.getByTitle("TemporaryEmailAttachments").rootFolder();
                                console.log(folder);
                                const filename=encodeURIComponent(attchment.fileName);
                                const addResult: IFileAddResult = await this.props.sp.web
                                    .getFolderByServerRelativePath(folder.ServerRelativeUrl)
                                    .files.addUsingPath(
                                        folder.ServerRelativeUrl + "/" + filename, attchment.content, { Overwrite: true });
                                console.log(addResult);
                                console.log(addResult.data.ServerRelativeUrl);
                                //window.location.pathname = addResult.data.ServerRelativeUrl;
                                const opts=`width=${window.innerWidth-100},height=${window.innerHeight-100},top=${window.screenTop+50},left=${window.screenLeft+50},toolbar=0,location=0`;
                                window.open(addResult.data.ServerRelativeUrl,filename,opts);

                            }}>
                            {att.fileName}
                        </DefaultButton>
                    </StackItem>

                )
            });
        return (<Panel
            isLightDismiss={true}
            isOpen={this.state.isOpen}
            headerText={this.props.headerText}
            isBlocking={false}
            closeButtonAriaLabel='Close'
            onDismiss={(e) => { this.setState({ isOpen: false }); }}
            type={PanelType.large}

        >
    
            <table>
                <tbody>
                    <tr>
                        <td style={{ fontWeight: 'bold' }}>From:</td>
                        <td>{this.props.message.getFileData().senderName} &lt;{this.props.message.getFileData().senderEmail}&gt;</td>
                    </tr>
                    <tr>
                        <td style={{ fontWeight: 'bold' }}>Sent on:</td>
                        <td>{this.props.message.getFileData().messageDeliveryTime}</td>
                    </tr> <tr>
                        <td style={{ fontWeight: 'bold' }}>To:</td>
                        <td>{toEmails}</td>
                    </tr>
                    <tr>
                        <td style={{ fontWeight: 'bold' }}>CC:</td>
                        <td>{ccEmails}</td>
                    </tr>
                    <tr>
                        <td style={{ fontWeight: 'bold' }}>Subject:</td>
                        <td>{this.props.message.getFileData().subject}</td>
                    </tr>
                    <tr>
                        <td style={{ fontWeight: 'bold' }}>Attachments:</td>
                        <td>
                            <Stack horizontal tokens={itemAlignmentsStackTokens} disableShrink={false} wrap={true}>

                                {att}

                            </Stack>

                        </td>
                    </tr>
                </tbody>
            </table>

            <TextField value={this.props.message.getFileData().body} multiline={true} autoAdjustHeight={true} />
            <DetailsList
                items={this.props.message.getFileData().attachments.filter(item => {

                    return !item.attachmentHidden
                })}
                columns={columns}
            ></DetailsList>
        </Panel>);
    }
}