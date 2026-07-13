/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import {
    Dropdown,
    IDropdownOption,
    TextField,
    PrimaryButton,
    DetailsList,
    Spinner,
    MessageBar,
    MessageBarType,
    DefaultButton,
    // IconButton,
    Stack
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SearchService } from "../service/SearchService";
import { columnsConfig } from "../constants/ColumnsConfig";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";


interface ISearchState {
    columns: IListColumn[];
    selectedColumn: string;
    rows: { columnKey: string, query: string }[]; // Dynamic rows with selected column and query
    results: ISearchResults[];
    loading: boolean;
    error: string | null;
}

interface ISearchProps {
    context: WebPartContext;
    listName: string;
}

class SearchComponent extends React.Component<ISearchProps, ISearchState> {
    private searchService: SearchService;

    constructor(props: ISearchProps) {
        super(props);
        this.searchService = new SearchService(props.context, props.listName);

        this.state = {
            columns: [],
            selectedColumn: "",
            rows: [{ columnKey: "", query: "" }], // Initially one row
            results: [],
            loading: false,
            error: null,
        };
    }

    async componentDidMount() {
        try {
            const columns = await this.searchService.loadColumns();
            this.setState({ columns });
        } catch (error) {
            this.setState({ error: error.message });
        }
    }

    // handleSearch = async () => {
    //     const { rows, columns } = this.state;
    //     if (rows.some(row => !row.columnKey || !row.query)) return;

    //     this.setState({ loading: true, error: null, results: [] });

    //     try {
    //         let results: ISearchResults[] = [];
    //         for (const row of rows) {
    //             const selectedColumnInfo = columns.find(col => col.key === row.columnKey);
    //             if (!selectedColumnInfo) throw new Error("Selected column not found");

    //             let searchResults: ISearchResults[] = [];
    //             if (selectedColumnInfo.fieldType === "Lookup") {
    //                 searchResults = await this.searchService.handleLookupSearch(selectedColumnInfo, row.query);
    //             } else {
    //                 searchResults = await this.searchService.handleStandardSearch([{ columnName: row.columnKey, query: row.query }]);
    //             }

    //             results = [...results, ...searchResults];
    //         }

    //         this.setState({ results, loading: false });
    //     } catch (error) {
    //         this.setState({ error: error.message, loading: false });
    //     }
    // };
    handleSearch = async () => {
        const { rows, columns } = this.state;
        if (rows.some(row => !row.columnKey || !row.query)) return;

        this.setState({ loading: true, error: null, results: [] });

        try {
            // Collect all standard filters (non-lookup)
            const standardFilters = rows
                .filter(row => {
                    const column = columns.find(col => col.key === row.columnKey);
                    return column?.fieldType !== "Lookup"; // Exclude Lookup columns
                })
                .map(row => ({ columnName: row.columnKey, query: row.query }));

            // Collect Lookup filters
            const lookupFilters = rows
                .filter(row => {
                    const column = columns.find(col => col.key === row.columnKey);
                    return column?.fieldType === "Lookup";
                });

            let results: ISearchResults[] = [];

            // Handle standard filters (AND logic)
            if (standardFilters.length > 0) {
                const standardResults = await this.searchService.handleStandardSearch(standardFilters);
                results = standardResults;
            }

            // Handle Lookup filters (AND logic)
            if (lookupFilters.length > 0) {
                for (const row of lookupFilters) {
                    const column = columns.find(col => col.key === row.columnKey);
                    if (!column) throw new Error("Column not found");
                    const lookupResults = await this.searchService.handleLookupSearch(column, row.query);
                    // Merge results only if there are existing results
                    results = results.length > 0
                        ? results.filter(item => lookupResults.some(lr => lr.Id === item.Id))
                        : lookupResults;
                }
            }

            this.setState({ results, loading: false });
        } catch (error) {
            this.setState({ error: error.message, loading: false });
        }
    };

    handleColumnChange = (index: number, _event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        const { rows } = this.state;
        rows[index].columnKey = option?.key.toString() || "";
        this.setState({ rows: [...rows] });
    };

    handleQueryChange = (index: number, _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        const { rows } = this.state;
        rows[index].query = newValue || "";
        this.setState({ rows: [...rows] });
    };

    addRow = () => {
        const { rows } = this.state;
        this.setState({ rows: [...rows, { columnKey: "", query: "" }] });
    };

    removeRow = (index: number) => {
        const { rows } = this.state;
        rows.splice(index, 1);
        this.setState({ rows: [...rows] });
    };

    handleClear = () => {
        this.setState({
            rows: [{ columnKey: "", query: "" }],
            results: [],
            loading: false,
            error: null,
            selectedColumn: ""
        });
    };
    render() {
        const { columns, rows, results, loading, error } = this.state;
        const columnsWithAction = columnsConfig.map((column) => {
            if (column.key === "approvalDetails") {
                return {
                    ...column,
                    onRender: (item: ISearchResults) => {
                        // Construct the SharePoint list URL with filters
                        const siteUrl = this.props.context.pageContext.web.absoluteUrl;
                        //const listName = encodeURIComponent(this.props.listName);
                        const listName = "ASDMSapprovals"
                        const titleFilter = encodeURIComponent(item.Title);

                        //const listUrl = `${siteUrl}/Lists/${listName}/AllItems.aspx?FilterField1=LinkTitle&FilterValue1=${titleFilter}&FilterType1=Computed`;
                        const listUrl = `${siteUrl}/Lists/${listName}/AllItems.aspx?FilterField1=LinkTitle&FilterValue1=${titleFilter}&FilterType1=Computed&FilterField2=Status&FilterValue2=Pending`;
                        return (
                            <PrimaryButton
                                text="Approval Details"
                                onClick={() => window.open(listUrl, "_blank")}
                            />
                        );
                    },
                };
            }
            return column;
        });

        return (
            <div className="search-container" style={{ padding: 20 }}>
                {/* Dynamic Rows */}
                {rows.map((row, index) => (
                    <Stack horizontal tokens={{ childrenGap: 10 }} style={{ margin: 14 }} verticalAlign="center" key={index}>
                        <Dropdown
                            placeholder="Select Option"
                            options={columns.map(c => ({ key: c.key, text: c.text }))}
                            selectedKey={row.columnKey}
                            onChange={(e, option) => this.handleColumnChange(index, e, option)}
                            styles={{ dropdown: { width: 200 } }}
                        />

                        <TextField
                            // placeholder="Enter Customer Name or Customer ID"
                            placeholder={
                                row.columnKey === "Title"
                                    ? "Enter Customer Name"
                                    : row.columnKey === "field_1"
                                        ? "Enter Customer ID"
                                        : "Enter Customer Name or Customer ID"
                            }
                            value={row.query}
                            onChange={(e, newValue) => this.handleQueryChange(index, e, newValue)}
                            disabled={!row.columnKey}
                            styles={{ root: { width: 250 } }}
                        />

                        {/* <IconButton
                            iconProps={{ iconName: "Add" }}{() => this.removeRow(index)}
                            disabled={rows.length <= 1} // Disable Remove for the last row
                        /> */}
                        {/* <PrimaryButton

                            text="Remove"
                            onClick={() => this.removeRow(index)}
                            disabled={rows.length <= 1} // Disable Remove for the last row
                        // styles={{ root: { marginTop: 15 } }}
                        /> */}

                        <PrimaryButton
                            text="Search"
                            onClick={this.handleSearch}
                            disabled={rows.some(row => !row.columnKey || !row.query)} // Disable if any row is incomplete
                        // styles={{ root: { marginTop: 15 } }}
                        />
                        <DefaultButton
                            text="Clear"
                            onClick={this.handleClear}
                            styles={{ root: { marginLeft: 5 } }}
                        />
                    </Stack>
                ))}

                {/* Search Button */}


                {/* Loading Indicator */}
                {loading && <Spinner label="Searching..." />}

                {/* Error Message */}
                {error && (
                    <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 15 } }}>
                        {error}
                    </MessageBar>
                )}

                {/* Search Results */}
                {results.length > 0 && (
                    <DetailsList
                        items={results}
                        columns={columnsWithAction}
                        isHeaderVisible={true}
                        styles={{ root: { marginTop: 20 } }}
                    />
                )}

                {/* No Results Message */}
                {!loading && !error && results.length === 0 && (
                    <MessageBar styles={{ root: { marginTop: 15 } }}>
                        No results found
                    </MessageBar>
                )}
            </div>
        );
    }
}

export default SearchComponent;