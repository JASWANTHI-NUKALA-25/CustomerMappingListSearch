
/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";


export class SearchService {
    private context: WebPartContext;
    private listName: string;

    constructor(context: WebPartContext, listName: string) {
        this.context = context;
        this.listName = listName;
    }

    // async loadColumns(): Promise<IListColumn[]> {
    //     try {
    //         const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;
    //         const response = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
    //         const data = await response.json();

    //         const columns = data.value.map((field: any) => ({
    //             key: field.InternalName,
    //             text: field.Title,
    //             fieldType: field.TypeAsString,
    //             lookupListId: field.LookupList,
    //             lookupField: field.LookupField
    //         }));

    //         // Manually add the CreatedBy (Author) field
    //         columns.push({
    //             key: "CreatedBy",
    //             text: "Created By",
    //             fieldType: "User",
    //             lookupListId: null,
    //             lookupField: null
    //         });

    //         return columns;
    //     } catch (err) {
    //         throw new Error(`Failed to load columns: ${err.message}`);
    //     }
    // }
//Update

async loadColumns(): Promise<IListColumn[]> {
        try {
            
            const columns = [
                
                {
                    key: "Title",
                    text: "Customer Name",
                    fieldType: "Text",
                    lookupListId: undefined,
                    lookupField: undefined
                },

                {
                    key: "field_1",
                    text: "Customer ID",
                    fieldType: "Text",
                    lookupListId: undefined,
                    lookupField: undefined
                },
                
                // {
                //     key: "DocType",
                //     text: "Document Type",
                //     fieldType: "Lookup",
                //     lookupListId: "f8b1c0d2-3e4f-4a5b-8c6d-7e8f9a0b1c2d", // Replace with actual DocType list ID
                //     lookupField: "Title"
                // },

                // {
                //     key: "Status",
                //     text: "Status",
                //     fieldType: "Choice",
                //     lookupListId: undefined,
                //     lookupField: undefined
                // },
                // {
                //     key: "RoleMemberEmail",
                //     text: "Rolemember",
                //     fieldType: "Note",
                //     lookupListId: undefined,
                //     lookupField: undefined
                // },
                // {
                //     key: "Supplier",
                //     text: "Supplier",
                //     fieldType: "Note",
                //     lookupListId: undefined,
                //     lookupField: undefined
                // },
                // {
                //     key: "BU",//Inter name
                //     text: "Business Unit",//Display Name
                //     fieldType: "Text",
                //     lookupListId: undefined,
                //     lookupField: undefined
                // },
                // {
                //     key: "PartNumber",
                //     text: "Part Number",
                //     fieldType: "Note",
                //     lookupListId: undefined,
                //     lookupField: undefined
                // },
                // {
                //     key: "CreatedBy",
                //     text: "Created By",
                //     fieldType: "User",
                //     lookupListId: undefined,
                //     lookupField: undefined
                // },
            ]

            // Manually add the CreatedBy (Author) field
            // columns.push({
            //     key: "CreatedBy",
            //     text: "Created By",
            //     fieldType: "User",
            //     lookupListId: null,
            //     lookupField: null
            // });

            return columns;
        } catch (err) {
            throw new Error(`Failed to load columns: ${err.message}`);
        }
    }
   
    async handleLookupSearch(columnInfo: IListColumn, query: string): Promise<ISearchResults[]> {

        
        // if (!columnInfo.lookupListId) throw new Error("Lookup list ID not found for this column");

        // const lookupField = columnInfo.lookupField || 'Title';
        // const lookupListUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=startswith(${lookupField}, '${query}')&$select=Id`;

        // const lookupResponse = await this.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        // const lookupData = await lookupResponse.json();

        if (!columnInfo.lookupListId) throw new Error("Lookup list ID not found");

        const lookupField = columnInfo.lookupField || 'Title';
        const lookupListUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=substringof('${query}', ${lookupField})`; // Use substringof for partial matches

        const lookupResponse = await this.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        const lookupData = await lookupResponse.json();

        if (!lookupData.value || lookupData.value.length === 0) {
            throw new Error("No matching items found in the lookup list");
        }

        const lookupIds = lookupData.value.map((item: any) => item.Id).join(",");
        //const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType&$filter=${columnInfo.key}/Id in (${lookupIds})`;
        const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Title,Docstype/Title,Status,BU,Supplier,Sequence,PartNum,RoleMemberEmail,${columnInfo.key}/Title&$expand=Docstype,${columnInfo.key}&$filter=${columnInfo.key}/Id eq ${lookupIds}`;

        const response = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        return data.value || [];
    }

    
//Update
    // async handleStandardSearch(filters: { columnName: string; query: string }[]): Promise<ISearchResults[]> {
    //     try {
    //         //const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType`;
    //         // const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType`;
    //       const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,field_1,field_2,field_3,field_4,field_5,field_6,field_7,field_8,field_9,field_10,field_11,field_12,field_13,field_14`;
    //         const response: SPHttpClientResponse = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
    //         const data = await response.json();
    //         const filteredResults = data.value.filter((item: any) =>
    //             filters.every(filter => {
    //                 let fieldValue = "";

    //                 // Handle CreatedBy (Author) field separately
    //                 if (filter.columnName === "CreatedBy") {
    //                     fieldValue = item.Author?.Title?.toLowerCase() || ""; // Access the Author/Title field
    //                 } else {
    //                     fieldValue = item[filter.columnName]?.toString().toLowerCase() || "";
    //                 }

    //                 const queryValue = filter.query.toLowerCase();
    //                 return fieldValue.includes(queryValue); // Partial match
    //             })
    //         );

    //         console.log("Filtered Results:", filteredResults);
            

    //         return filteredResults;

    //     } catch (error) {
    //         throw new Error(`Search failed: ${error.message}`);
    //     }
    // }

async handleStandardSearch(filters: { columnName: string; query: string }[]): Promise<ISearchResults[]> {
    try {
        const baseFields = [
            'Title', 'field_1', 'field_2', 'field_3', 'field_4', 'field_5',
            'field_6', 'field_7', 'field_8', 'field_9', 'field_10',
            'field_11', 'field_12', 'field_13', 'field_14'
        ];

        const selectSet = new Set(baseFields);
        const expandSet = new Set<string>();

        for (const f of filters) {
            if (f.columnName === 'CreatedBy') {
                selectSet.add('Author/Title');
                expandSet.add('Author');
            } else {
                selectSet.add(f.columnName);
            }
        }

        const select = Array.from(selectSet).join(',');
        const expand = Array.from(expandSet).join(',');

        const top = 4999; // fetch up to 4999 items
        let searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=${select}&$top=${top}`;
        if (expand) {
            searchUrl += `&$expand=${expand}`;
        }

        console.log('Search request URL:', searchUrl);

        const response: SPHttpClientResponse = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        const filteredResults = (data.value || []).filter((item: any) =>
            filters.every(filter => {
                let fieldValue = '';

                if (filter.columnName === 'CreatedBy') {
                    fieldValue = item.Author?.Title?.toString().toLowerCase() || '';
                } else {
                    fieldValue = (item[filter.columnName] ?? '').toString().toLowerCase();
                }

                const queryValue = filter.query.toLowerCase().trim();
                return fieldValue.includes(queryValue);
            })
        );

        return filteredResults;
    } catch (error) {
        throw new Error(`Search failed: ${error.message}`);
    }
}


}