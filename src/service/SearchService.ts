
/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";

// Internal names that are system-managed metadata, not part of the customer mapping data itself.
const SYSTEM_FIELD_KEYS = new Set([
    "ID", "Created", "Modified", "Author", "Editor", "ContentType", "Attachments",
    "Edit", "LinkTitle", "LinkTitleNoMenu", "ItemChildCount", "FolderChildCount",
    "AppAuthor", "AppEditor", "_ComplianceAssetId", "WorkflowVersion", "GUID", "Order",
    "FileSystemObjectType", "FileRef", "FileDirRef", "MetaInfo", "_ModerationComments",
    "_ModerationStatus", "InstanceID", "OData__ColorTag", "OData__CopySource",
    "SortBehavior", "ParentVersionString", "ParentLeafName", "OData__UIVersionString",
    "SyncClientId", "TemplateUrl", "owshiddenversion"
]);

const CUSTOMER_ID_DISPLAY_NAME = "erp customer id";

export class SearchService {
    private context: WebPartContext;
    private listName: string;
    private siteUrl: string;

    constructor(context: WebPartContext, listName: string, siteUrl?: string) {
        this.context = context;
        this.listName = listName;
        this.siteUrl = siteUrl || context.pageContext.web.absoluteUrl;
    }

    private listFields: IListColumn[] | undefined;

    /** Discovers the list's real (non-system) columns directly from SharePoint - no hardcoded field names. */
    async getListFields(): Promise<IListColumn[]> {
        if (this.listFields) return this.listFields;

        const url = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=InternalName,Title,TypeAsString`;
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) {
            throw new Error(`Failed to load list columns (status ${response.status})`);
        }
        const data = await response.json();

        const fields: IListColumn[] = (data.value || [])
            .filter((f: any) => !SYSTEM_FIELD_KEYS.has(f.InternalName))
            .map((f: any) => ({
                key: f.InternalName,
                text: f.Title,
                fieldType: f.TypeAsString,
                lookupListId: undefined,
                lookupField: undefined
            }));

        // eslint-disable-next-line no-console
        console.log(`CustomerMapping list fields for '${this.listName}'`, fields);

        this.listFields = fields;
        return fields;
    }

    /** Columns offered in the search dropdown: Customer Name and Customer ID only. */
    async loadColumns(): Promise<IListColumn[]> {
        const fields = await this.getListFields();
        return fields.filter(f => f.key === "Title" || f.text.trim().toLowerCase() === CUSTOMER_ID_DISPLAY_NAME);
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
    private requestDigest: string | undefined;

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

    /** SPHttpClient only auto-manages the digest for the *current* web, so when the list lives on a
     *  different site than the one hosting the webpart, a fresh digest for that site must be fetched explicitly. */
    private async getRequestDigest(): Promise<string> {
        if (this.requestDigest) return this.requestDigest;

        const url = `${this.siteUrl}/_api/contextinfo`;
        const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: { "Accept": "application/json;odata=verbose" }
        });
        if (!response.ok) {
            throw new Error(`Failed to get request digest (status ${response.status})`);
        }
        const data = await response.json();
        const digest: string | undefined = data?.d?.GetContextWebInformation?.FormDigestValue;
        if (!digest) {
            throw new Error("Could not resolve request digest for the target site");
        }
        this.requestDigest = digest;
        return digest;
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
        const [entityType, digest] = await Promise.all([this.getEntityTypeFullName(), this.getRequestDigest()]);
        const url = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items`;
        const body = JSON.stringify({ "__metadata": { type: entityType }, ...item });

        const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: {
                "Accept": "application/json;odata=verbose",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
                "X-RequestDigest": digest
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
