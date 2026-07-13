// /* eslint-disable @typescript-eslint/explicit-function-return-type */
// /* eslint-disable @rushstack/no-new-null */
// import * as React from "react";
// import {
//     Dropdown,
//     IDropdownOption,
//     TextField,
//     PrimaryButton,
//     DetailsList,
//     Spinner,
//     MessageBar,
//     MessageBarType,
//     IconButton,
//     Stack
// } from "@fluentui/react";
// import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { SearchService } from "../service/SearchService";
// import { columnsConfig } from "../constants/ColumnsConfig";
// import { IListColumn } from "../interfaces/IListColumn";
// import { ISearchResults } from "../interfaces/ISearchResults.ts";


// interface ISearchState {
//     columns: IListColumn[];
//     selectedColumn: string;
//     rows: { columnKey: string, query: string }[]; // Dynamic rows with selected column and query
//     results: ISearchResults[];
//     loading: boolean;
//     error: string | null;
// }

// interface ISearchProps {
//     context: WebPartContext;
//     listName: string;
// }

// class BackUpSearchComponent extends React.Component<ISearchProps, ISearchState> {
//     private searchService: SearchService;

//     constructor(props: ISearchProps) {
//         super(props);
//         this.searchService = new SearchService(props.context, props.listName);

//         this.state = {
//             columns: [],
//             selectedColumn: "",
//             rows: [{ columnKey: "", query: "" }], // Initially one row
//             results: [],
//             loading: false,
//             error: null,
//         };
//     }

//     async componentDidMount() {
//         try {
//             const columns = await this.searchService.loadColumns();
//             this.setState({ columns });
//         } catch (error) {
//             this.setState({ error: error.message });
//         }
//     }

//     handleSearch = async () => {
//         const { rows, columns } = this.state;
//         if (rows.some(row => !row.columnKey || !row.query)) return; // Ensure all rows are filled

//         this.setState({ loading: true, error: null, results: [] });

//         try {
//             const results = [];
//             for (const row of rows) {
//                 const selectedColumnInfo = columns.find(col => col.key === row.columnKey);
//                 if (!selectedColumnInfo) throw new Error("Selected column not found");

//                 const searchResults = selectedColumnInfo.fieldType === "Lookup"
//                     ? await this.searchService.handleLookupSearch(selectedColumnInfo, row.query)
//                     : await this.searchService.handleStandardSearch(row.columnKey, row.query);

//                 results.push(...searchResults);
//             }

//             this.setState({ results, loading: false });
//         } catch (error) {
//             this.setState({ error: error.message, loading: false });
//         }
//     };

//     handleColumnChange = (index: number, _event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
//         const { rows } = this.state;
//         rows[index].columnKey = option?.key.toString() || "";
//         this.setState({ rows: [...rows] });
//     };

//     handleQueryChange = (index: number, _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
//         const { rows } = this.state;
//         rows[index].query = newValue || "";
//         this.setState({ rows: [...rows] });
//     };

//     addRow = () => {
//         const { rows } = this.state;
//         this.setState({ rows: [...rows, { columnKey: "", query: "" }] });
//     };

//     removeRow = (index: number) => {
//         const { rows } = this.state;
//         rows.splice(index, 1);
//         this.setState({ rows: [...rows] });
//     };

//     render() {
//         const { columns, rows, results, loading, error } = this.state;

//         return (
//             <div className="search-container" style={{ padding: 20 }}>
//                 {/* Dynamic Rows */}
//                 {rows.map((row, index) => (
//                     <Stack horizontal tokens={{ childrenGap: 10 }} style={{ margin: 14 }} verticalAlign="center" key={index}>
//                         <Dropdown
//                             placeholder="Select Column"
//                             options={columns.map(c => ({ key: c.key, text: c.text }))}
//                             selectedKey={row.columnKey}
//                             onChange={(e, option) => this.handleColumnChange(index, e, option)}
//                             styles={{ dropdown: { width: 200 } }}
//                         />

//                         <TextField
//                             placeholder="Enter search value"
//                             value={row.query}
//                             onChange={(e, newValue) => this.handleQueryChange(index, e, newValue)}
//                             disabled={!row.columnKey}
//                             styles={{ root: { width: 250 } }}
//                         />

//                         <IconButton
//                             iconProps={{ iconName: "Add" }}
//                             title="Add"
//                             onClick={this.addRow}
//                         />

//                         <IconButton
//                             iconProps={{ iconName: "Remove" }}
//                             title="Remove"
//                             onClick={() => this.removeRow(index)}
//                             disabled={rows.length <= 1} // Disable Remove for the last row
//                         />
//                     </Stack>
//                 ))}

//                 {/* Search Button */}
//                 <PrimaryButton
//                     text="Search"
//                     onClick={this.handleSearch}
//                     disabled={rows.some(row => !row.columnKey || !row.query)} // Disable if any row is incomplete
//                     styles={{ root: { marginTop: 15 } }}
//                 />

//                 {/* Loading Indicator */}
//                 {loading && <Spinner label="Searching..." />}

//                 {/* Error Message */}
//                 {error && (
//                     <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 15 } }}>
//                         {error}
//                     </MessageBar>
//                 )}

//                 {/* Search Results */}
//                 {results.length > 0 && (
//                     <DetailsList
//                         items={results}
//                         columns={columnsConfig}
//                         isHeaderVisible={true}
//                         styles={{ root: { marginTop: 20 } }}
//                     />
//                 )}

//                 {/* No Results Message */}
//                 {!loading && !error && results.length === 0 && (
//                     <MessageBar styles={{ root: { marginTop: 15 } }}>
//                         No results found
//                     </MessageBar>
//                 )}
//             </div>
//         );
//     }
// }

// export default BackUpSearchComponent;