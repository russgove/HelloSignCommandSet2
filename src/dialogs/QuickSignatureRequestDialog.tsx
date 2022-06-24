import { BaseDialog, Dialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ActionButton, DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { DialogContent, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { TextField } from "office-ui-fabric-react/lib/TextField";
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { QuickSignatureRequest, Signer } from '../model/model';


interface IQuickSignatureRequestContentProps {

  // add props here to create signatire request 
  message: string;
  title: string;
  close: () => void;
  aadHttpClient: AadHttpClient;
  azureFunctionUrl: string;
  azureWakeUpUrl: string;
  itemId: number;
  userEmail: string;
}
interface IQuickSignatureRequestContentState {
  sigReq: QuickSignatureRequest;
  isSaving: boolean;
  errors: Array<string>;


}

class QuickSignatureRequestContent extends React.Component<IQuickSignatureRequestContentProps, IQuickSignatureRequestContentState> {
  constructor(props: any) {
    super(props);
    let sigReq = new QuickSignatureRequest();
    sigReq.signers = [];
    sigReq.ccs = [];
    this.state = {
      sigReq: sigReq,
      isSaving: false,
      errors: []
    };
    this.WakeUpAzure();
  }
  public WakeUpAzure(): Promise<any> {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    return this.props.aadHttpClient.fetch(this.props.azureWakeUpUrl,
      AadHttpClient.configurations.v1,
      {
        method: "GET",
        headers: requestHeaders,
      })
      .then((response: HttpClientResponse) => {
        if (response.status === 200) {

        } else {

          console.log(`Error Waking up azure, reponse follows`);
          console.log(response);
          let errors = this.state.errors;
          errors.push("Error Waking up Azure Function");
          this.setState((current) => ({ ...current, errors: errors }));
          return null;
        }

      }).catch((err) => {
        debugger;
        console.log(`Error waking up azure http Reponse follows`);
        console.log(err);
        let errors = this.state.errors;
        errors.push("HTTP Error Waking up Azure Function");
        this.setState((current) => ({ ...current, errors: errors }));

        debugger;
        return null;
      });
  }
  public AcceptAsIs(): Promise<any> {
    this.setState((current) => ({ ...current, isSaving: true }));
    const body: string = JSON.stringify({
      'UserEmail': this.props.userEmail,
      "ItemId": this.props.itemId,
      "Comments": this.state.sigReq.name

    });
    const requestHeaders: Headers = new Headers();
    debugger;
    requestHeaders.append('Content-type', 'application/json');
    return this.props.aadHttpClient.fetch(this.props.azureFunctionUrl,
      AadHttpClient.configurations.v1,
      {
        method: "POST",
        body: body,
        headers: requestHeaders,
      })
      .then((response: HttpClientResponse) => {
        this.setState((current) => ({ ...current, isSaving: false }));

        if (response.status === 200) {
          this.props.close();
        } else {
          debugger;
          console.log(`Error acepting http Reponse follows`);
          console.log(response);
          let errors = this.state.errors;
          errors.push(`Error-- Code: ${response.status} ${response.statusText}`);
          this.setState((current) => ({ ...current, errors: errors }));
          return null;
        }

      }).catch((err) => {
        let errors = this.state.errors;
        errors.push(`an error occurred`);
        errors.push(err.message);
        this.setState((current) => ({ ...current, errors: errors }));

        debugger;
        return null;
      });
  }


  public render(): JSX.Element {
    debugger;
    return <DialogContent
      title={this.props.title}
      onDismiss={this.props.close}
      showCloseButton={true}
    >
      <div>
        {this.state.errors.map((error, i) => {
          console.log("Entered");
          // Return the errors
          return (<MessageBar messageBarType={MessageBarType.error}  >{error} </MessageBar>);
        })}
      </div>
      <Label>{this.props.message}</Label>
      <TextField label="Name" defaultValue={this.state.sigReq.name}
        onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
          let tempSigreq = this.state.sigReq;
          tempSigreq.name = newValue;
          this.setState((current) => ({ ...current, sigReq: tempSigreq }));
        }}
      >
      </TextField>

      <DetailsList
        items={this.state.sigReq.signers}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        selectionMode={SelectionMode.none}
        columns={[
          {
            key: "name", name: "Name", fieldName: "name", minWidth: 20, maxWidth: 300,
            onRender: (item, idx) =>
              <TextField defaultValue={item.name} placeholder="enter name" errorMessage={item.name && item.name!=''?null:'A signer name is required'}

                onChange={(element, value) => {
                  debugger;
                  let tempSigreq = this.state.sigReq;
                  tempSigreq.signers[idx].name = value;
                  this.setState((current) => ({ ...current, sigReq: tempSigreq }));
                }}>
              </TextField>

          },
          {
            key: "emailAddress", name: "Email address", fieldName: "emailAddress", minWidth: 30, maxWidth: 300,
            onRender: (item: Signer, idx) =>
              <TextField defaultValue={item.emailAddress} placeholder="email address" errorMessage={item.emailAddress.match(
                /^[a-zA-Z0-9.!#$%&’*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/) ? null : 'A valid email address is required'}

                onChange={(element, value) => {
                  debugger;
                  let tempSigreq = this.state.sigReq;
                  tempSigreq.signers[idx].emailAddress = value;
                  this.setState((current) => ({ ...current, sigReq: tempSigreq }));
                }}>
              </TextField>

          }
        ]}
      >
      </DetailsList>
      <ActionButton onClick={(e) => {
        let tempSigreq = this.state.sigReq;
        tempSigreq.signers.push({ emailAddress: '', name: '', pin: '', order: tempSigreq.signers.length + 1 });
        this.setState((current) => ({ ...current, sigReq: tempSigreq }));
      }}>
        + Add another signer
           </ActionButton>
           <ActionButton onClick={(e) => {
        let tempSigreq = this.state.sigReq;
        tempSigreq.signers.push({ emailAddress: '', name: '', pin: '', order: tempSigreq.signers.length + 1 });
        this.setState((current) => ({ ...current, sigReq: tempSigreq }));
      }}>
        + Add me as signer
           </ActionButton>


      <DialogFooter>
        <DefaultButton disabled={this.state.isSaving} onClick={() => this.props.close()}   >Cancel</DefaultButton>
        <PrimaryButton disabled={this.state.isSaving || this.state.sigReq.name === ""} onClick={() => this.AcceptAsIs()}   >OK</PrimaryButton>
      </DialogFooter>
      <ProgressIndicator progressHidden={!this.state.isSaving} />
    </DialogContent>;
  }
}
export default class QuickSignatureRequestDialog extends BaseDialog {
  public message: string;
  public title: string;
  public aadHttpClient: AadHttpClient;
  public azureFunctionUrl: string;
  public azureWakeUpUrl: string;
  public itemId: number;
  public userEmail: string;



  public render(): void {

    ReactDOM.render(<QuickSignatureRequestContent
      azureFunctionUrl={this.azureFunctionUrl}
      azureWakeUpUrl={this.azureWakeUpUrl}
      itemId={this.itemId}
      userEmail={this.userEmail}
      message={this.message}
      close={this.close}
      aadHttpClient={this.aadHttpClient}
      title={this.title}
    />, this.domElement);


  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: true
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
  }


}