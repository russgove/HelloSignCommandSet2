import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
} from "@microsoft/sp-listview-extensibility";
import { spfi, SPFx } from "@pnp/sp";
import HelloSign from "hellosign-embedded";
import { Dialog } from "@microsoft/sp-dialog";
import {
  AadHttpClient,
  HttpClientResponse,
  AadHttpClientConfiguration,
} from "@microsoft/sp-http";
import * as strings from "HelloSignCommandSetStrings";
import CheckStatusDialog from "../../dialogs/CheckStatus";

import QuickSignatureRequestDialog from "../../dialogs/QuickSignatureRequestDialog";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloSignCommandSetProperties {
  hellosignFunctionBaseUrl: string;
  hellosignFunctionClientID: string;
  helloSignClientId: string;
  supportedFileTypes: string;
  adminWebUrl: string;
}

const LOG_SOURCE: string = "HelloSignCommandSet";

export default class HelloSignCommandSet extends BaseListViewCommandSet<IHelloSignCommandSetProperties> {
  private aadHttpClient: AadHttpClient;
  private isProcessing: boolean = false; //prevent double clicking
  public onInit(): Promise<void> {
    debugger;
    return super.onInit().then((_) => {
      debugger;
      const sp = spfi().using(SPFx(this.context));
console.log(`hellosignFunctionBaseUrl=${this.properties.hellosignFunctionBaseUrl}`);
console.log(`hellosignFunctionClientID=${this.properties.hellosignFunctionClientID}`);
 
      return this.context.aadHttpClientFactory
        .getClient(this.properties.hellosignFunctionClientID)
        //.getClient("becb3efa-2875-4eda-8fb8-40d03f7cb4e7")
        .then((client): void => {
          // connect to the API
          debugger;
          this.aadHttpClient = client;
          this.wakeUpService();
        })
        .catch((e) => {
          debugger;
        });
    });
  }
  private wakeUpService() {
    debugger;

    this.aadHttpClient
      .get(
        `${this.properties.hellosignFunctionBaseUrl}/api/WakeUp`,
        AadHttpClient.configurations.v1
      )
      .then((response: HttpClientResponse) => {
        console.log("service is awake");
      })
      .catch((e) => {
        console.error(e);
        alert("failed to wake up service");
      });
  }
  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const quicksign: Command = this.tryGetCommand("COMMAND_QUICKSIGN");
    const embeddedsign: Command = this.tryGetCommand(
      "COMMAND_EMBEDDEDSIGNATURE"
    );
    const checkStatus: Command = this.tryGetCommand("COMMAND_CHECKSTATUS");
    if (quicksign) {
      // This command should be hidden unless exactly one row is selected.
      quicksign.visible = false; //event.selectedRows.length === 1;
    }
    if (embeddedsign) {
      // This command should be hidden unless exactly one row is selected
      // and its not a folder
      // and its a supported file type
      // and its not too deeply nested (400 chhsrs max in path)
      embeddedsign.visible =
        event.selectedRows.length === 1 &&
        event.selectedRows[0].getValueByName("FSObjType") !== "1" &&
        this.properties.supportedFileTypes.indexOf(
          event.selectedRows[0].getValueByName("File_x0020_Type")
        ) !== -1;
    }
    if (checkStatus) {
      // This command should be hidden unless exactly one row is selected
      // and its not a folder
      // and its a supported file type
      // and its not too deeply nested (400 chhsrs max in path)
      checkStatus.visible =
        event.selectedRows.length === 1 &&
        event.selectedRows[0].getValueByName("FSObjType") !== "1" &&
        this.properties.supportedFileTypes.indexOf(
          event.selectedRows[0].getValueByName("File_x0020_Type")
        ) !== -1;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_QUICKSIGN":
        this.cmdQuickSign(event);
        break;

      case "COMMAND_EMBEDDEDSIGNATURE":
        this.cmdEmbeddedSign(event);
        break;
      case "COMMAND_CHECKSTATUS":
        this.cmdChekStatus(event);

        break;
      default:
        throw new Error("Unknown command");
    }
  }
  private cmdQuickSign(event: IListViewCommandSetExecuteEventParameters) {
    const quickSignDlg: QuickSignatureRequestDialog =
      new QuickSignatureRequestDialog();
    quickSignDlg.title = `Send document  ${event.selectedRows[0].getValueByName(
      "FileLeafRef"
    )} to HelloSign`;
    quickSignDlg.itemId = event.selectedRows[0].getValueByName("ID");
    quickSignDlg.aadHttpClient = this.aadHttpClient;
    quickSignDlg.message =
      "Enter details belos and click the send button to intiate the signature";
    quickSignDlg.azureFunctionUrl = `${this.properties.hellosignFunctionBaseUrl}/api/EmbeddedSignatureRequest`;
    quickSignDlg.azureWakeUpUrl = `${this.properties.hellosignFunctionBaseUrl}/api/WakeUp`; // dialog call this to 'wake up' the azure function
    quickSignDlg.userEmail = this.context.pageContext.user.email;
    quickSignDlg.show();
  }
  private cmdChekStatus(event: IListViewCommandSetExecuteEventParameters) {
    const dialog: CheckStatusDialog = new CheckStatusDialog();
    dialog.title = `CHECK STATUS`;
    dialog.aadHttpClient = this.aadHttpClient;
    dialog.userEmail = this.context.pageContext.user.email;
    dialog.webServerRelativeUrl =
      this.context.pageContext.web.serverRelativeUrl;
    dialog.helloSignClientId = this.properties.helloSignClientId;
    dialog.adminWebUrl = this.properties.adminWebUrl;
    dialog.signatureRequestListName = "SignatureRequests";
    dialog.documentUniqueId = event.selectedRows[0].getValueByName("UniqueId");
    dialog.hellosignFunctionBaseUrl = this.properties.hellosignFunctionBaseUrl;
    dialog.context = this.context;
    dialog.show();
  }
  private cmdEmbeddedSign(event: IListViewCommandSetExecuteEventParameters) {
    console.log(`in embedded signg isProcessing=${this.isProcessing}`);
    if (!this.isProcessing) {
      this.isProcessing = true;
      this.CreateEmbeddedSignatureRequest(event);
    }
  }
  public CreateEmbeddedSignatureRequest(
    event: IListViewCommandSetExecuteEventParameters
  ): Promise<any> {
    debugger;
    const requestHeaders: Headers = new Headers();

    //nocors when test8ing local nope, add  "Host": {
    // "CORS": "*"
    // } in local-settings.json
    //requestHeaders.append('mode', 'no-cors');
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      userEmail: this.context.pageContext.user.email,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      fileServerRelativeUrl: event.selectedRows[0].getValueByName("FileRef"),
      fileName: event.selectedRows[0].getValueByName("FileLeafRef"),
      siteUrl: window.location.origin,
      libraryName: this.context.pageContext.list.title,
      documentUniqueId: event.selectedRows[0].getValueByName("UniqueId"),
    });
    console.log(body);
    return this.aadHttpClient
      .fetch(
        `${this.properties.hellosignFunctionBaseUrl}/api/CreateUnclaimedDraft`,
        AadHttpClient.configurations.v1,
        {
          method: "POST",
          headers: requestHeaders,
          body: body,
        }
      )
      .then(async (response: HttpClientResponse) => {
        this.isProcessing = false; //just preventing a doubleclick
        if (response.status === 200) {
          var resp = await response.json();
          if (resp.errorName) {
            console.error(resp);
            alert(`HelloSign Error : ${resp.message}`);
            return;
          }
          const client = new HelloSign();
          client.open(resp.claimUrl, {
            clientId: this.properties.helloSignClientId,
            debug: true,
            //   skipDomainVerification: true
          });
        } 
        else { // response.status !== 200
          alert("HTTP Error in CreateEmbeddedSignatureRequest");
          console.error(
            `HTTP Error in CreateEmbeddedSignatureRequest, reponse follows`
          );
          console.error(response);
          return null;
        }
      })
      .catch((err) => {
        this.isProcessing = false; //just preventing a doubleclick
       debugger;
        alert("Error in CreateEmbeddedSignatureRequest");
        console.log(`Error in CreateEmbeddedSignatureRequest Reponse follows`);
        console.log(err);

        return null;
      });
  }
}
