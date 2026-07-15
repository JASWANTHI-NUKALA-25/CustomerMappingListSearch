
/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";

// Matches the CustomerMapping1 list schema exactly (confirmed via List Settings > columns).
const LIST_FIELDS: IListColumn[] = [
    { key: "Title", text: "Customer Name", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_1", text: "Customer ID", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_2", text: "Region", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_3", text: "Americas Lead", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_4", text: "Americas Director", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_5", text: "Americas G/KAM", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_6", text: "Americas Sales Agent", fieldType: "Text", lookupListId: undefined, lookupField: undefined },
    { key: "field_7", text: "EMEA Lead", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_8", text: "EMEA Director", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_9", text: "EMEA G/KAM", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_10", text: "EMEA Sales Agent", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_11", text: "APAC Lead", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_12", text: "APAC Director", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_13", text: "APAC G/KAM", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
    { key: "field_14", text: "APAC Sales Agent", fieldType: "Number", lookupListId: undefined, lookupField: undefined },
];

export class SearchService {
    private context: WebPartContext;
    private listName: string;
    private siteUrl: string;

    constructor(context: WebPartContext, listName: string, siteUrl?: string) {
        this.context = context;
        this.listName = listName;
        this.siteUrl = siteUrl || context.pageContext.web.absoluteUrl;
    }

    /** Returns the CustomerMapping1 list's known columns (hardcoded - see LIST_FIELDS above). */
    async getListFields(): Promise<IListColumn[]> {
        return LIST_FIELDS;
    }

    /** Columns offered in the search dropdown: Customer Name and Customer ID only. */
    async loadColumns(): Promise<IListColumn[]> {
        const fields = await this.getListFields();
        return fields.filter(f => f.key === "Title" || f.key === "field_1");
    }

    async getFieldTypeMap(): Promise<Record<string, string>> {
        const fields = await this.getListFields();
        const map: Record<string, string> = {};
        fields.forEach(f => { map[f.key] = f.fieldType; });
        return map;
    }

    async handleLookupSearch(columnInfo: IListColumn, query: string): Promise<ISearchResults[]> {
        if (!columnInfo.lookupListId) throw new Error("Lookup list ID not found");

        const lookupField = columnInfo.lookupField || "Title";
        const lookupListUrl = `${this.siteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=substringof('${query}', ${lookupField})`;

        const lookupResponse = await this.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        const lookupData = await lookupResponse.json();

        if (!lookupData.value || lookupData.value.length === 0) {
            throw new Error("No matching items found in the lookup list");
        }

        const lookupIds = lookupData.value.map((item: any) => item.Id).join(",");
        const fields = await this.getListFields();
        const select = fields.map(f => f.key).join(",");
        const searchUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=${select}&$filter=${columnInfo.key}/Id eq ${lookupIds}`;

        const response = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        return data.value || [];
    }

    async handleStandardSearch(filters: { columnName: string; query: string }[]): Promise<ISearchResults[]> {
        try {
            const fields = await this.getListFields();
            const selectSet = new Set(fields.map(f => f.key));

            for (const f of filters) {
                selectSet.add(f.columnName);
            }

            const select = Array.from(selectSet).join(",");
            const top = 4999; // fetch up to 4999 items

            const searchUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=${select}&$top=${top}`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
            const data = await response.json();

            const filteredResults = (data.value || []).filter((item: any) =>
                filters.every(filter => {
                    const fieldValue = (item[filter.columnName] ?? "").toString().toLowerCase();
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

    private async getEntityTypeFullName(): Promise<string> {
        if (this.entityTypeFullName) return this.entityTypeFullName;

        const url = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')?$select=ListItemEntityTypeFullName`;
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
        const url = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items`;
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
            // eslint-disable-next-line no-console
            console.error("CustomerMapping item creation failed", { item, entityType, errorText });
            const message = this.extractErrorMessage(errorText) || `Request failed with status ${response.status}`;
            throw new Error(`${message} (payload: ${JSON.stringify(item)})`);
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
