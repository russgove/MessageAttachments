/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { BaseComponentContext } from "@microsoft/sp-component-base";

import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { TextField, } from 'office-ui-fabric-react/lib/TextField';
import { Stack, StackItem, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { SPFI } from "@pnp/sp";
import MsgReader, { FieldsData } from "@kenjiuno/msgreader";
export interface IAttachmentPanelV2Props {

    sp: SPFI;
    context: BaseComponentContext;
    itemId: number;
    headerText: string;

}
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IFileAddResult } from '@pnp/sp/files';
export interface IAttachmentPanelV2State {

    isOpen: boolean;
    loading: boolean;
    frameUrl?: string;
    message?: MsgReader;

}
export default class AttachmentPanel extends React.Component<IAttachmentPanelV2Props, IAttachmentPanelV2State> {



    // public constructor(msg: MsgReader) {
    //     super({});
    //     this.message = msg;
    //     this.messageFields = this.message.getFileData();
    //     this.messageAttachments = this.messageFields.attachments;
    // }
    public constructor(props: IAttachmentPanelV2Props) {
        super(props);
        this.state = { isOpen: true, loading: true };
    }
    public componentDidMount(): void {
        this.props.sp.web.lists
            .getById(this.props.context.pageContext.list.id.toString())
            .items.getById(this.props.itemId)
            .file()
            .then(async (fileInfo) => {
                const searchParams = new URLSearchParams(window.location.href);
                const parent = encodeURIComponent(searchParams.get("id"));
                const frameUrl = `${window.location.origin}${window.location.pathname
                    }?id=${encodeURIComponent(fileInfo.ServerRelativeUrl)
                        .split("_")
                        .join("%5F")
                        .split(".")
                        .join("%2E")}&parent=${parent}`;
                console.log(frameUrl);
                this.setState((current) => ({ ...current, frameUrl: frameUrl }));
                debugger;
                const url = fileInfo.ServerRelativeUrl;
                const buffer: ArrayBuffer = await this.props.sp.web
                    .getFileByServerRelativePath(url)
                    .getBuffer();
                debugger;
                const tmpmsg = new MsgReader(buffer);
                this.setState((current) => ({ ...current, message: tmpmsg, loading: false }));

            })
            .catch((e) => {
                debugger;
            });
    }


    public render(): React.ReactElement<{}> {
        const itemAlignmentsStackTokens: IStackTokens = {
            childrenGap: 5,
            padding: 10,
        };
        debugger;
        let att;
        if (this.state.message) {
            att = this.state.message.getFileData().attachments
                .filter(item => {

                    return !item.attachmentHidden
                }).map(att => {
                    return (
                        <StackItem>
                            <DefaultButton className='ms-bgColor-themeLight'
                                onClick={async (e) => {
                                    debugger;
                                    console.log(att);
                                    const attchment = this.state.message.getAttachment(att);
                                    const folder = await this.props.sp.web.lists
                                    .getByTitle("TemporaryEmailAttachments")
                                    .rootFolder() .catch(e => {
                                        alert(e);
                                    });
                                    if (!folder){return;}
                                    console.log(folder);
                                    const filename = encodeURIComponent(attchment.fileName);
                                    const addResult: IFileAddResult | void = await this.props.sp.web
                                        .getFolderByServerRelativePath(folder.ServerRelativeUrl)
                                        .files.addUsingPath(
                                            folder.ServerRelativeUrl + "/" + filename, attchment.content, { Overwrite: true })
                                        .catch(e => {
                                            alert(e);
                                        })
                                        ;
                                    if (addResult) {
                                       debugger;
                                        console.log(addResult.data.ServerRelativeUrl);
                                        const newUrl=`${window.location.origin}${this.props.context.pageContext.web.serverRelativeUrl}/TemporaryEmailAttachments/Forms/AllItems.aspx?id=${encodeURIComponent(addResult.data.ServerRelativeUrl)}&parent=${encodeURIComponent(folder.ServerRelativeUrl)}`;
                                        console.log(newUrl);
                                        //window.location.pathname = addResult.data.ServerRelativeUrl;

                                        
                                        const opts = `width=${window.innerWidth - 100},height=${window.innerHeight - 100},top=${window.screenTop + 50},left=${window.screenLeft + 50},toolbar=0,location=0`;
                                        window.open(newUrl, filename, opts);
                                    }

                                }}>
                                {att.name}
                            </DefaultButton>
                        </StackItem>

                    )
                });
        }

        return (<Panel
            isLightDismiss={true}
            isOpen={this.state.isOpen}
            headerText={this.props.headerText}
            isBlocking={false}
            closeButtonAriaLabel='Close'
            onDismiss={(e) => { this.setState({ isOpen: false }); }}
            type={PanelType.large}

        >
            {
                this.state.loading && <Spinner label="Getting list data" size={SpinnerSize.medium} />
            }
            <Stack horizontal tokens={itemAlignmentsStackTokens} disableShrink={false} wrap={true}>
                {att}
            </Stack>
            <iframe src={this.state.frameUrl} height={window.innerHeight * .7} width={window.innerWidth * .7} />
            {/* <table>
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

            <TextField value={this.props.message.getFileData().body} multiline={true} autoAdjustHeight={true} /> */}
            {/* <DetailsList
                items={this.props.message.getFileData().attachments.filter(item => {

                    return !item.attachmentHidden
                })}
                columns={columns}
            ></DetailsList> */}
        </Panel>);
    }
}