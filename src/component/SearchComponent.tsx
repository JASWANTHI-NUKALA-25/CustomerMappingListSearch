/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import {
    Dropdown,
    IDropdownOption,
    TextField,
    PrimaryButton,
    DetailsList,
    IDetailsListProps,
    DetailsListLayoutMode,
    Spinner,
    MessageBar,
    MessageBarType,
    DefaultButton,
    IconButton,
    Stack,
    Text,
    Separator,
    ScrollablePane,
    ScrollbarVisibility,
    Sticky,
    StickyPositionType
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SearchService } from "../service/SearchService";
import { columnsConfig } from "../constants/ColumnsConfig";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";
import ExcelUploadComponent from "./ExcelUploadComponent";


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
    handleQueryKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
        const { rows } = this.state;
        if (event.key === "Enter" && !rows.some(row => !row.columnKey || !row.query)) {
            this.handleSearch();
        }
    };

    onRenderDetailsHeader: IDetailsListProps["onRenderDetailsHeader"] = (props, defaultRender) => {
        if (!props || !defaultRender) return null;
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                {defaultRender({ ...props })}
            </Sticky>
        );
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

        const cardStyle: React.CSSProperties = {
            maxWidth: 1200,
            margin: "0 auto",
            padding: 24,
            background: "#ffffff",
            borderRadius: 4,
            boxShadow: "0 1.6px 3.6px rgba(0,0,0,0.1), 0 0.3px 0.9px rgba(0,0,0,0.08)"
        };

        return (
            <div style={{ padding: 20 }}>
                <div style={cardStyle}>
                    <Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 4 } }}>
                        Customer Mapping List Search
                    </Text>
                    <Text variant="medium" block styles={{ root: { color: "#605e5c", marginBottom: 16 } }}>
                        Search existing customer mapping records or bulk-upload new ones from Excel.
                    </Text>

                    {/* Excel Upload */}
                    <ExcelUploadComponent
                        context={this.props.context}
                        listName={this.props.listName}
                    />

                    <Separator styles={{ root: { margin: "16px 0" } }} />

                    {/* Search Filters */}
                    <Stack tokens={{ childrenGap: 12 }}>
                        {rows.map((row, index) => {
                            const selectedColumnInfo = columns.find(c => c.key === row.columnKey);
                            return (
                                <Stack
                                    horizontal
                                    wrap
                                    tokens={{ childrenGap: 10 }}
                                    verticalAlign="center"
                                    key={index}
                                >
                                    <Dropdown
                                        placeholder="Select Option"
                                        options={columns.map(c => ({ key: c.key, text: c.text }))}
                                        selectedKey={row.columnKey}
                                        onChange={(e, option) => this.handleColumnChange(index, e, option)}
                                        styles={{ dropdown: { width: 220 } }}
                                    />

                                    <TextField
                                        placeholder={
                                            selectedColumnInfo
                                                ? `Enter ${selectedColumnInfo.text}`
                                                : "Select a field first"
                                        }
                                        value={row.query}
                                        onChange={(e, newValue) => this.handleQueryChange(index, e, newValue)}
                                        onKeyPress={this.handleQueryKeyPress}
                                        disabled={!row.columnKey}
                                        styles={{ root: { width: 260 } }}
                                    />

                                    <IconButton
                                        iconProps={{ iconName: "Remove" }}
                                        title="Remove this filter"
                                        ariaLabel="Remove this filter"
                                        onClick={() => this.removeRow(index)}
                                        disabled={rows.length <= 1}
                                    />

                                    {index === rows.length - 1 && (
                                        <IconButton
                                            iconProps={{ iconName: "Add" }}
                                            title="Add another filter"
                                            ariaLabel="Add another filter"
                                            onClick={this.addRow}
                                        />
                                    )}
                                </Stack>
                            );
                        })}

                        <Stack horizontal tokens={{ childrenGap: 10 }}>
                            <PrimaryButton
                                text="Search"
                                onClick={this.handleSearch}
                                disabled={rows.some(row => !row.columnKey || !row.query)}
                            />
                            <DefaultButton
                                text="Clear"
                                onClick={this.handleClear}
                            />
                        </Stack>
                    </Stack>

                    {/* Loading Indicator */}
                    {loading && <Spinner label="Searching..." styles={{ root: { marginTop: 16 } }} />}

                    {/* Error Message */}
                    {error && (
                        <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginTop: 16 } }}>
                            {error}
                        </MessageBar>
                    )}

                    {/* Search Results */}
                    {results.length > 0 && (
                        <div style={{ marginTop: 20, position: "relative", height: "60vh" }}>
                            <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                                <DetailsList
                                    items={results}
                                    columns={columnsWithAction}
                                    isHeaderVisible={true}
                                    layoutMode={DetailsListLayoutMode.fixedColumns}
                                    onRenderDetailsHeader={this.onRenderDetailsHeader}
                                />
                            </ScrollablePane>
                        </div>
                    )}

                    {/* No Results Message */}
                    {!loading && !error && results.length === 0 && (
                        <MessageBar styles={{ root: { marginTop: 16 } }}>
                            No results found
                        </MessageBar>
                    )}
                </div>
            </div>
        );
    }
}

export default SearchComponent;