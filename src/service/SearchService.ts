
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

            const columns: IListColumn[] = [
                { key: "Title", text: "Customer Name", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
                { key: "field_1", text: "Customer ID", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
            ];

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

    private entityTypeFullName: string | undefined;
    private fieldTypeMap: Record<string, string> | undefined;

    async getFieldTypeMap(): Promise<Record<string, string>> {
        if (this.fieldTypeMap) return this.fieldTypeMap;

        const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/fields?$select=InternalName,TypeAsString`;
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) {
            throw new Error(`Failed to resolve list field types (status ${response.status})`);
        }
        const data = await response.json();
        const map: Record<string, string> = {};
        (data.value || []).forEach((field: any) => {
            map[field.InternalName] = field.TypeAsString;
        });

        this.fieldTypeMap = map;
        return map;
    }

    private async getEntityTypeFullName(): Promise<string> {
        if (this.entityTypeFullName) return this.entityTypeFullName;

        const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')?$select=ListItemEntityTypeFullName`;
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) {
            throw new Error(`Failed to resolve list metadata (status ${response.status})`);
        }
        const data = await response.json();
        const entityType: string | undefined = data.ListItemEntityTypeFullName;
        if (!entityType) {
            throw new Error("Could not resolve the list's entity type");
        }
        this.entityTypeFullName = entityType;
        return entityType;
    }

    private extractErrorMessage(errorText: string): string {
        try {
            const parsed = JSON.parse(errorText);
            return parsed?.error?.message?.value || parsed?.["odata.error"]?.message?.value || errorText;
        } catch (e) {
            return errorText;
        }
    }

    async createItem(item: Record<string, string | number | boolean>): Promise<void> {
        const entityType = await this.getEntityTypeFullName();
        const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items`;
        const body = JSON.stringify({ "__metadata": { type: entityType }, ...item });

        const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: {
                "Accept": "application/json;odata=verbose",
                "Content-type": "application/json;odata=verbose",
                "odata-version": ""
            },
            body
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(this.extractErrorMessage(errorText) || `Request failed with status ${response.status}`);
        }
    }

    async bulkCreateItems(entries: { row: number; data: Record<string, string | number | boolean> }[]): Promise<{ success: number; failed: { row: number; error: string }[] }> {
        const failed: { row: number; error: string }[] = [];
        let success = 0;
        const concurrency = 5;

        for (let i = 0; i < entries.length; i += concurrency) {
            const chunk = entries.slice(i, i + concurrency);
            const outcomes = await Promise.all(
                chunk.map(entry =>
                    this.createItem(entry.data)
                        .then(() => ({ ok: true, row: entry.row, error: "" }))
                        .catch(err => ({ ok: false, row: entry.row, error: err.message }))
                )
            );

            outcomes.forEach(outcome => {
                if (outcome.ok) {
                    success++;
                } else {
                    failed.push({ row: outcome.row, error: outcome.error });
                }
            });
        }

        return { success, failed };
    }

}