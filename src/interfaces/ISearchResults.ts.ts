/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ISearchResults {
    Title: string;
    DocsType?: { Docs: string };
    Status: string;
    BU: string;
    Supplier:string;
    Sequence:string;
    PartNum:string;
    RoleMemberEmail:String;
    [key: string]: any;
}
