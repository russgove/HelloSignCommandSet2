import { BaseDialog, Dialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

import { format } from "date-fns";

import { DetailsList, SelectionMode, Selection } from "office-ui-fabric-react/lib/DetailsList";
import { spfi, SPFx } from "@pnp/sp";
import { DialogContent } from "office-ui-fabric-react/lib/Dialog";
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';

import { Link } from "office-ui-fabric-react/lib/Link";
import { MessageBar, MessageBarType, } from "office-ui-fabric-react/lib/MessageBar";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";

import { TextField } from "office-ui-fabric-react/lib/TextField";
import * as React from "react";
import * as ReactDOM from "react-dom";

import {
  SentRequest,
  SignatureRequestFromHS,
  SignatureItemFromHS,
} from "../model/model";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { BaseComponentContext } from "@microsoft/sp-component-base";

interface ICheckStatusContentProps {
  // add props here to create signatire request

  title: string;
  close: () => void;
  aadHttpClient: AadHttpClient;

  helloSignClientId: string;
  userEmail: string;

  adminWebUrl: string;
  signatureRequestListName: string;
  documentUniqueId: string;
  hellosignFunctionBaseUrl: string;
  context: BaseComponentContext;
}
interface ICheckStatusContentState {
  errors: Array<string>;
  signatureRequests: Array<SentRequest>;
  signatureRequestFromHS: SignatureRequestFromHS;
  selectionDetails: string;
}

class CheckStatusContent extends React.Component<
  ICheckStatusContentProps,
  ICheckStatusContentState
> {
  private _selection: Selection;
  constructor(props: any) {
    super(props);
    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });
    this.state = {
      errors: [],
      signatureRequests: [],
      signatureRequestFromHS: null, selectionDetails: "",
    };
  }
  private _getSelectionDetails(): string { //TODO:  why do i need this??
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0]);
      default:
        return `${selectionCount} items selected`;
    }
  }
  public componentDidMount() {

    const admin = spfi(this.props.adminWebUrl).using(SPFx(this.context));
    admin.web.lists
      .getByTitle(this.props.signatureRequestListName)
      .items.filter(`DocumentUniqueId eq '${this.props.documentUniqueId}'`)()

      .then((res) => {
        const requests = res.map((r) => {
          return {
            SignatureRequestId: r.SignatureRequestId,
            DocumentUniqueId: r.DocumentUniqueId,
            SenderEmail: r.SenderEmail,
            Title: r.Title,
            Status: r.Status,
            Created: r.Created,
          };
        });
        this.setState((current) => ({
          ...current,
          signatureRequests: requests,
        }));
      })
      .catch((res) => {
        debugger;
      });
  }
  public cancelSignatureRequest(SignatureRequestId: string) {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      signatureRequestId: SignatureRequestId,
    });
    debugger;
    return this.props.aadHttpClient
      .fetch(
        `${this.props.hellosignFunctionBaseUrl}/api/CancelSignatureRequest`,
        AadHttpClient.configurations.v1,
        {
          method: "POST",
          headers: requestHeaders,
          body: body,
        }
      )
      .then(async (response: HttpClientResponse) => {
        debugger;
        if (response.status === 200) {
          alert("Canceled!");
          this.setState((current) => ({
            ...current,
            signatureRequestFromHS: null,
          }));
          this.componentDidMount();
        } else {
          debugger;
          alert("HTTP Error in CancelSignatureRequest");
          console.error(
            `HTTP Error in CancelSignatureRequest, reponse follows`
          );
          console.error(response);
          return null;
        }
      })
      .catch((err) => {
        debugger;
        alert("Error in CreateEmbeddedSignatureRequest");
        console.log(`Error in CreateEmbeddedSignatureRequest Reponse follows`);
        console.log(err);
      });
  }
  public getStatus(SignatureRequestId: string) {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      signatureRequestId: SignatureRequestId,
    });
    return this.props.aadHttpClient
      .fetch(
        `${this.props.hellosignFunctionBaseUrl}/api/GetSignatureRequest`,
        AadHttpClient.configurations.v1,
        {
          method: "POST",
          headers: requestHeaders,
          body: body,
        }
      )
      .then(async (response: HttpClientResponse) => {
        debugger;
        if (response.status === 200) {
debugger;
          var result: SignatureRequestFromHS = await response.json();
          this.setState((current) => ({
            ...current,
            signatureRequestFromHS: result,
          }));
        } else {
          debugger;
          alert("HTTP Error in getStatus");
          console.error(
            `HTTP Error in getStatus, reponse follows`
          );
          console.error(response);
          return null;
        }
      })
      .catch((err) => {
        debugger;
        alert("Error in getStatus");
        console.log(`Error in getStatus Reponse follows`);
        console.log(err);
      });
  }
  public renderActionTime(value: string): string {

    var result = (value !== "0001-01-01T00:00:00") ?
      format(new Date(value), "dd-MMM-yyyy@hh:mm:ss a") : "-";

    return result;

  }
  public render(): JSX.Element {
    const _items: ICommandBarItemProps[] = [

      {
        key: 'View',
        text: 'View on HelloSign',
        iconProps: { iconName: 'View' },
        href: this.state.signatureRequestFromHS ? this.state.signatureRequestFromHS.detailsUrl : null, target: "_blank"
      },
      {
        key: 'cancel',
        disabled: this.state.signatureRequestFromHS ? this.state.signatureRequestFromHS.isComplete === true : false,
        text: 'Cancel Signature Request',
        iconProps: { iconName: 'Cancel' },
        onClick: () => {
          debugger;
          this.cancelSignatureRequest(this.state.signatureRequestFromHS.signatureRequestId);


        }
      },
      // {
      //   disabled: this._selection.count !== 1,
      //   key: 'edit',
      //   text: 'Edit Signatory Email',
      //   iconProps: { iconName: 'Edit' },
      //   onClick: () => { debugger; console.log(this._selection.count); console.log('Download'); },
      // },
    ];
    return (
      <DialogContent
        title={"Signature Requests"}
        onDismiss={this.props.close}
        showCloseButton={true}
      >
        <div>
          {this.state.errors.map((error, i) => {
            console.log("Entered");
            // Return the errors
            return (
              <MessageBar messageBarType={MessageBarType.error}>
                {error}{" "}
              </MessageBar>
            );
          })}
        </div>


        {this.state.signatureRequestFromHS !== null && (
          <Panel
            type={PanelType.large} headerText="HelloSign Status"
            isOpen={this.state.signatureRequestFromHS !== null}
            onDismiss={(e) => {

              this.setState((current) => ({
                ...current,
                signatureRequestFromHS: null,
              }));
            }}
          >
            <CommandBar items={_items} />


            <TextField
              label="Title"
              value={this.state.signatureRequestFromHS.title}
              borderless={true}
            />
            <TextField
              label="Subject"
              value={this.state.signatureRequestFromHS.subject}
              borderless={true}
            />
            <TextField
              label="Message"
              value={this.state.signatureRequestFromHS.message}
              borderless={true}
            />
            <TextField
              label="Requester"
              value={this.state.signatureRequestFromHS.requesterEmailAddress}
              borderless={true}
            />
            <DetailsList selection={this._selection}
              items={this.state.signatureRequestFromHS.signatures}
              // layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.single}
              columns={[
                {
                  key: "signerName",
                  name: "Name",
                  fieldName: "signerName",
                  minWidth: 200,

                  isResizable: true,
                },
                {
                  key: "signerEmailAddress",
                  name: "Email",
                  fieldName: "signerEmailAddress",
                  minWidth: 250,

                  isResizable: true,
                },
                {
                  key: "statusCode",
                  name: "Status",
                  fieldName: "statusCode",
                  minWidth: 120,
                  isResizable: true,
                },
                {
                  key: "signedAt",
                  name: "Date Signed",
                  fieldName: "signedAt",
                  minWidth: 200,
                  isResizable: true,
                  onRender: (item: SignatureItemFromHS) => {
                    return this.renderActionTime(item.signedAt);

                  },
                },
                {
                  key: "name",
                  name: "Last Reminded",
                  fieldName: "lastRemindedAt",
                  minWidth: 200,
                  isResizable: true,
                  onRender: (item: SignatureItemFromHS) => {
                    return this.renderActionTime(item.lastRemindedAt);

                  },
                },
                {
                  key: "name",
                  name: "Last Viewed",
                  fieldName: "lastViewedAt",
                  minWidth: 200
                  ,
                  isResizable: true,
                  onRender: (item: SignatureItemFromHS) => {
                    return this.renderActionTime(item.lastViewedAt);

                  },
                }

              ]}
            ></DetailsList>
          </Panel>
        )}
        <DetailsList
          items={this.state.signatureRequests}
          // layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.none}
          columns={[
            {
              key: "Title",
              name: "Title",
              fieldName: "Title",
              minWidth: 200,
              isResizable: true,
            },
            {
              key: "SenderEmail",
              name: "SenderEmail",
              fieldName: "SenderEmail",
              minWidth: 250,
              maxWidth: 400,
              isResizable: true,
            },
            {
              key: "Created",
              name: "Created",
              fieldName: "Created",
              minWidth: 150,
              isResizable: true,
              onRender: (item: SentRequest) => {
                return this.renderActionTime(item.Created);

              },
            },
            {
              key: "Status",
              name: "Status",
              fieldName: "Status",
              minWidth: 60,
              maxWidth: 60,
              isResizable: true,
            },

            // {
            //   key: "name", name: "Docid", fieldName: "DocumentUniqueId", minWidth: 100, maxWidth: 100,isResizable:true

            // },
            {
              key: "SignatureRequestId",
              name: "Details",
              fieldName: "SignatureRequestId",
              minWidth: 50,
              maxWidth: 50,
              onRender: (item?: SentRequest, index?: number): any => {
                if (item.Status === `Canceled`) {
                  return (<div></div>);
                } else return (<Link onClick={() => {
                  this.getStatus(item.SignatureRequestId);
                }}
                >View</Link>)
              }

              ,







            },
          ]}
        ></DetailsList>
      </DialogContent>
    );
  }
}
export default class CheckStatusDialog extends BaseDialog {

  public title: string;
  public aadHttpClient: AadHttpClient;
  public webServerRelativeUrl: string;
  public userEmail: string;
  public helloSignClientId: string;
  public adminWebUrl: string;
  public signatureRequestListName: string;
  public documentUniqueId: string;
  public hellosignFunctionBaseUrl: string;
  public context: BaseComponentContext;

  public render(): void {
    ReactDOM.render(
      <CheckStatusContent
        adminWebUrl={this.adminWebUrl}
        hellosignFunctionBaseUrl={this.hellosignFunctionBaseUrl}
        documentUniqueId={this.documentUniqueId}
        signatureRequestListName={this.signatureRequestListName}
        userEmail={this.userEmail}
        context={this.context}
        close={this.close}
        aadHttpClient={this.aadHttpClient}
        title={this.title}
        helloSignClientId={this.helloSignClientId}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: true,
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}
