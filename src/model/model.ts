export class Signer {
    public emailAddress: string;
    public name: string;
    public order: number;
    public pin: string;
}
export class QuickSignatureRequest {
    public ccs?: Array<string>;
    public fileUrls?: Array<string>;
    public signers: Array<Signer>;
    public allowReassign?: boolean;
    public allowDecline?: boolean;
    public name?: string;
    public order?: number;
    public pin?: string;
}
export class SentRequest {
    public DocumentUniqueId: string;
    public SignatureRequestId: string;
    public SenderEmail: string;
    public Title: string;
    public Status: string;
    public Created: string;
}
export class SignatureRequestFromHS{
    public   signatureRequestId:string;
    public   title:string;
    public   subject:string;
    public  message:string;
    public  isComplete:boolean;
    public  detailsUrl:string;
    public  requesterEmailAddress:string;
    public  ccEmailAddresses:Array<string>;
    public  responseData:Array<ResponseDataItemFromHS>;
    public  signatures:Array<SignatureItemFromHS>;
}
export class SignatureItemFromHS {
    public  signatureId: string;
    public  signerEmailAddress: string;
    public  signerName: string;
    public signerRole: string;
    public order?: number;
    public statusCode: string;
    public signedAt: string;
    public  lastViewedAt: string;
    public   lastRemindedAt: string;
    public   hasPin: boolean;
    public   declineReason?: string;
    public  error?: string;
    public reassignedBy?: string;
    public   reassignmentReason?: string;
}
export class ResponseDataItemFromHS {
    public   apiId: string;
    public  signatureId: string;
    public  name: string;
    public  value?: string;
    public  type: string;

}