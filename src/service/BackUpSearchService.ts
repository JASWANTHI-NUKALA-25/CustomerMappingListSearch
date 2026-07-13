/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";


export class BackUpSearchService {
    private context: WebPartContext;
    private listName: string;

    constructor(context: WebPartContext, listName: string) {
        this.context = context;
        this.listName = listName;
    }

    async loadColumns(): Promise<IListColumn[]> {
        try {
            const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;
            const response = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            const data = await response.json();

            return data.value.map((field: any) => ({
                key: field.InternalName,
                text: field.Title,
                fieldType: field.TypeAsString,
                lookupListId: field.LookupList,
                lookupField: field.LookupField
            }));
        } catch (err) {
            throw new Error(`Failed to load columns: ${err.message}`);
        }
    }

    async handleLookupSearch(columnInfo: IListColumn, query: string): Promise<ISearchResults[]> {
        if (!columnInfo.lookupListId) throw new Error("Lookup list ID not found for this column");

        const lookupField = columnInfo.lookupField || 'Title';
        const lookupListUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=${lookupField} eq '${query}'&$select=Id`;

        const lookupResponse = await this.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        const lookupData = await lookupResponse.json();

        if (!lookupData.value || lookupData.value.length === 0) {
            throw new Error("No matching items found in the lookup list");
        }

        const lookupId = lookupData.value[0].Id;
        const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Title,Docstype/Title,Status,BU,Supplier,Sequence,PartNum,${columnInfo.key}/Title&$expand=Docstype,${columnInfo.key}&$filter=${columnInfo.key}/Id eq ${lookupId}`;

        const response = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        return data.value || [];
    }

    async handleStandardSearch(columnName: string, query: string): Promise<ISearchResults[]> {
        const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,Docstype/Docs,Status,BU,Supplier,Sequence,PartNum&$expand=Docstype`;

        const response = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        return data.value.filter((item: any) =>
            item[columnName]?.toString().toLowerCase().includes(query.toLowerCase())
        );
    }
}
