export class QueryHelper {
    static getEncodedListUrl(baseUrl: string, listName: string): string {
        return `${baseUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`;
    }
}
