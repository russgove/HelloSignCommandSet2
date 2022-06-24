// import { BaseDialog, Dialog, IDialogConfiguration } from '@microsoft/sp-dialog';
// import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
// import HelloSign from 'hellosign-embedded';
// import { ActionButton, DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
// import { DetailsList, DetailsListLayoutMode, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
// import { DialogContent, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
// import { Label } from 'office-ui-fabric-react/lib/Label';
// import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
// import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
// import { TextField } from "office-ui-fabric-react/lib/TextField";
// import * as React from 'react';
// import * as ReactDOM from 'react-dom';

// import { QuickSignatureRequest, Signer } from '../model/model';



// interface IEmbeddedSignatureRequestContentProps {

//   // add props here to create signatire request 
//   message: string;
//   title: string;
//   close: () => void;
//   aadHttpClient: AadHttpClient;
//   webServerRelativeUrl: string;
//   azureCreateUnclaimedDraftUrl: string;
//   fileServerRelativeUrl: string;
//   siteUrl: string;
//   fileName: string;
//   helloSignClientId:string;
//   userEmail: string;
//   libraryName:string;
// }
// interface IEmbeddedSignatureRequestContentState {
//   sigReq: QuickSignatureRequest;
//   isSaving: boolean;
//   errors: Array<string>;


// }

// class EmbeddedSignatureRequestContent extends React.Component<IEmbeddedSignatureRequestContentProps, IEmbeddedSignatureRequestContentState> {
//   constructor(props: any) {
//     super(props);
//     let sigReq = new QuickSignatureRequest();
//     sigReq.signers = [];
//     sigReq.ccs = [];
//     this.state = {
//       sigReq: sigReq,
//       isSaving: false,
//       errors: []
//     };
//     this.CreateEmbeddedSignatureRequest();
//   }

//   public CreateEmbeddedSignatureRequest(): Promise<any> {

//     const requestHeaders: Headers = new Headers();

//     //nocors when test8ing local nope, add  "Host": {
//     // "CORS": "*"
//     // } in local-settings.json
//     //requestHeaders.append('mode', 'no-cors');
//     requestHeaders.append('Content-type', 'application/json');
//     const body: string = JSON.stringify({
//       'userEmail': this.props.userEmail,
//       'webServerRelativeUrl': this.props.webServerRelativeUrl,
//       "fileServerRelativeUrl": this.props.fileServerRelativeUrl,
//       "fileName": this.props.fileName,
//       "siteUrl": this.props.siteUrl,
//       "libraryName":this.props.libraryName

//     });
//     return this.props.aadHttpClient.fetch(this.props.azureCreateUnclaimedDraftUrl,
//       AadHttpClient.configurations.v1,
//       {
//         method: "POST",
//         headers: requestHeaders,
//         body: body
//       })
//       .then(async (response: HttpClientResponse) => {
//         console.log(response);
//         if (response.status === 200) {
//           debugger;
//           var resp = await response.json();
// console.dir(resp);
// if(resp.errorName){
//   console.error(resp);
//   alert(resp.message);
// return;
// }

//           const client = new HelloSign();
// console.log(`claimurl is ${resp.claimUrl}`);
//           client.open(resp.claimUrl, {
//             clientId: this.props.helloSignClientId,
//            debug:true,
//        //   skipDomainVerification: true
//           });
//           debugger;
//           this.props.close();

//         } else {

//           console.error(`HTTP Error in CreateEmbeddedSignatureRequest, reponse follows`);
//           console.error(response);
//           let errors = this.state.errors;
//           errors.push("Error inCreateEmbeddedSignatureRequest");
//           this.setState((current) => ({ ...current, errors: errors }));
//           return null;
//         }

//       }).catch((err) => {
//         debugger;
//         console.log(`Error in CreateEmbeddedSignatureRequest Reponse follows`);
//         console.log(err);
//         let errors = this.state.errors;
//         errors.push(err.message);
//         this.setState((current) => ({ ...current, errors: errors }));

//         debugger;
//         return null;
//       });
//   }


//   public render(): JSX.Element {
//     debugger;
//     return <DialogContent
//       title={this.props.title}
//       onDismiss={this.props.close}
//       showCloseButton={true}
//     >
//       <div>
//         {this.state.errors.map((error, i) => {
//           console.log("Entered");
//           // Return the errors
//           return (<MessageBar messageBarType={MessageBarType.error}  >{error} </MessageBar>);
//         })}
//       </div>

//       <Label>{this.props.message}</Label>
//       <TextField label="Name" defaultValue={this.state.sigReq.name}
//         onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
//           let tempSigreq = this.state.sigReq;
//           tempSigreq.name = newValue;
//           this.setState((current) => ({ ...current, sigReq: tempSigreq }));
//         }}
//       >
//       </TextField>

//       <DetailsList
//         items={this.state.sigReq.signers}
//         layoutMode={DetailsListLayoutMode.fixedColumns}
//         selectionMode={SelectionMode.none}
//         columns={[
//           {
//             key: "name", name: "Name", fieldName: "name", minWidth: 20, maxWidth: 300,
//             onRender: (item, idx) =>
//               <TextField defaultValue={item.name} placeholder="enter name" errorMessage={item.name && item.name != '' ? null : 'A signer name is required'}

//                 onChange={(element, value) => {
//                   debugger;
//                   let tempSigreq = this.state.sigReq;
//                   tempSigreq.signers[idx].name = value;
//                   this.setState((current) => ({ ...current, sigReq: tempSigreq }));
//                 }}>
//               </TextField>

//           },
//           {
//             key: "emailAddress", name: "Email address", fieldName: "emailAddress", minWidth: 30, maxWidth: 300,
//             onRender: (item: Signer, idx) =>
//               <TextField defaultValue={item.emailAddress} placeholder="email address" errorMessage={item.emailAddress.match(
//                 /^[a-zA-Z0-9.!#$%&’*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/) ? null : 'A valid email address is required'}

//                 onChange={(element, value) => {
//                   debugger;
//                   let tempSigreq = this.state.sigReq;
//                   tempSigreq.signers[idx].emailAddress = value;
//                   this.setState((current) => ({ ...current, sigReq: tempSigreq }));
//                 }}>
//               </TextField>

//           }
//         ]}
//       >
//       </DetailsList>
//       {/* <ActionButton onClick={(e) => {
//         let tempSigreq = this.state.sigReq;
//         tempSigreq.signers.push({ emailAddress: '', name: '', pin: '', order: tempSigreq.signers.length + 1 });
//         this.setState((current) => ({ ...current, sigReq: tempSigreq }));
//       }}>
//         + Add another signer
//            </ActionButton>
//       <ActionButton onClick={(e) => {
//         let tempSigreq = this.state.sigReq;
//         tempSigreq.signers.push({ emailAddress: '', name: '', pin: '', order: tempSigreq.signers.length + 1 });
//         this.setState((current) => ({ ...current, sigReq: tempSigreq }));
//       }}>
//         + Add me as signer
//            </ActionButton>
//  */}

//       <DialogFooter>
//         <DefaultButton disabled={this.state.isSaving} onClick={() => this.props.close()}   >Cancel</DefaultButton>
//         <PrimaryButton disabled={this.state.isSaving || this.state.sigReq.name === ""} onClick={() => { debugger; }}   >OK</PrimaryButton>
//       </DialogFooter>
//       <ProgressIndicator progressHidden={!this.state.isSaving} />
//     </DialogContent>;
//   }
// }
// export default class EmbeddedSignatureRequestDialog extends BaseDialog {
//   public message: string;
//   public title: string;
//   public aadHttpClient: AadHttpClient;
//   public webServerRelativeUrl: string;
//   public azureCreateUnclaimedDraftUrl: string;
// public libraryName :string;
//   public userEmail: string;
//   public fileServerRelativeUrl: string;
//   public siteUrl: string;
//   public fileName: string;
//   public helloSignClientId:string;



//   public render(): void {

//     ReactDOM.render(<EmbeddedSignatureRequestContent

//       azureCreateUnclaimedDraftUrl={this.azureCreateUnclaimedDraftUrl}

//       userEmail={this.userEmail}
//       message={this.message}
//       close={this.close}
//       aadHttpClient={this.aadHttpClient}
//       fileServerRelativeUrl={this.fileServerRelativeUrl}
//       fileName={this.fileName}
//       siteUrl={this.siteUrl}
//       libraryName={this.libraryName}
//       title={this.title}
//       webServerRelativeUrl={this.webServerRelativeUrl}
//       helloSignClientId={this.helloSignClientId}
//     />, this.domElement);


//   }

//   public getConfig(): IDialogConfiguration {
//     return {
//       isBlocking: true
//     };
//   }

//   protected onAfterClose(): void {
//     super.onAfterClose();
//     // Clean up the element for the next dialog
//     ReactDOM.unmountComponentAtNode(this.domElement);
//   }


// }