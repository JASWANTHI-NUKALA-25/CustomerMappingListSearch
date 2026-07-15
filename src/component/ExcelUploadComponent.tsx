/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import * as XLSX from "xlsx";
import {
    PrimaryButton,
    DefaultButton,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize,
    Stack,
    Text
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SearchService } from "../service/SearchService";

interface IExcelUploadProps {
    context: WebPartContext;
    listName: string;
    siteUrl: string;
    onUploadComplete?: () => void;
}

interface IExcelUploadState {
    uploading: boolean;
    message: { type: MessageBarType; text: string } | null;
}

const CUSTOMER_ID_DISPLAY_NAME = "customer id";

const isNumericFieldType = (fieldType?: string): boolean => fieldType === "Number" || fieldType === "Currency";

class ExcelUploadComponent extends React.Component<IExcelUploadProps, IExcelUploadState> {
    private searchService: SearchService;
    private fileInputRef = React.createRef<HTMLInputElement>();

    constructor(props: IExcelUploadProps) {
        super(props);
        this.searchService = new SearchService(props.context, props.listName, props.siteUrl);
        this.state = { uploading: false, message: null };
    }

    private triggerFileSelect = (): void => {
        if (this.fileInputRef.current) {
            this.fileInputRef.current.click();
        }
    };

    private handleDownloadTemplate = async (): Promise<void> => {
        try {
            const fields = await this.searchService.getListFields();
            const headers = fields.map(f => f.text);
            const worksheet = XLSX.utils.aoa_to_sheet([headers]);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, this.props.listName);
            XLSX.writeFile(workbook, `${this.props.listName}_Template.xlsx`);
        } catch (error) {
            this.setState({ message: { type: MessageBarType.error, text: error.message } });
        }
    };

    private handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
        const file = event.target.files && event.target.files[0];
        event.target.value = ""; // allow re-selecting the same file
        if (!file) return;

        this.setState({ uploading: true, message: null });

        try {
            const fields = await this.searchService.getListFields();
            const customerIdField = fields.find(f => f.text.trim().toLowerCase() === CUSTOMER_ID_DISPLAY_NAME);
            if (!customerIdField) {
                throw new Error(`Could not find a "${CUSTOMER_ID_DISPLAY_NAME}" column on the list.`);
            }

            const headerToKey: Record<string, string> = {};
            fields.forEach(f => { headerToKey[f.text.trim().toLowerCase()] = f.key; });
            const expectedHeaders = fields.map(f => f.text);
            const fieldTypes: Record<string, string> = {};
            fields.forEach(f => { fieldTypes[f.key] = f.fieldType; });

            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            if (!sheetName) throw new Error("The workbook does not contain any sheets.");

            const sheet = workbook.Sheets[sheetName];
            const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
            if (rows.length === 0) throw new Error("The uploaded file is empty.");

            const headerRow = (rows[0] || []).map((h: any) => (h || "").toString().trim());
            const missingHeaders = expectedHeaders.filter(
                expected => !headerRow.some(h => h.toLowerCase() === expected.toLowerCase())
            );

            if (missingHeaders.length > 0) {
                throw new Error(
                    `The file's columns don't match the SharePoint list. Missing column(s): ${missingHeaders.join(", ")}. Use "Download Template" to get the correct format.`
                );
            }

            const unrecognizedHeaders = headerRow.filter(h => h && !headerToKey[h.toLowerCase()]);

            const dataRows = rows.slice(1).filter(row => row.some(cell => (cell ?? "").toString().trim() !== ""));
            if (dataRows.length === 0) throw new Error("No data rows found in the uploaded file.");

            const entries: { row: number; data: Record<string, string | number> }[] = [];
            const rowErrors: string[] = [];

            dataRows.forEach((row, idx) => {
                const excelRowNumber = idx + 2; // header is row 1
                const item: Record<string, string | number> = {};
                let rowHasError = false;

                headerRow.forEach((header, colIdx) => {
                    const key = header ? headerToKey[header.toLowerCase()] : undefined;
                    if (!key) return;

                    const rawValue = (row[colIdx] ?? "").toString().trim();
                    if (rawValue === "") return;

                    if (isNumericFieldType(fieldTypes[key])) {
                        const numericValue = Number(rawValue.replace(/,/g, ""));
                        if (isNaN(numericValue)) {
                            rowErrors.push(`Row ${excelRowNumber}: "${header}" must be a number (got "${rawValue}").`);
                            rowHasError = true;
                        } else {
                            item[key] = numericValue;
                        }
                    } else {
                        item[key] = rawValue;
                    }
                });

                if (rowHasError) return;

                if (!item["Title"] || !item[customerIdField.key]) {
                    rowErrors.push(`Row ${excelRowNumber}: Customer Name and Customer ID are required.`);
                    return;
                }
                entries.push({ row: excelRowNumber, data: item });
            });

            if (entries.length === 0) {
                throw new Error(`No valid rows to upload. ${rowErrors.join(" ")}`);
            }

            const result = await this.searchService.bulkCreateItems(entries);
            const allErrors = [...rowErrors, ...result.failed.map(f => `Row ${f.row}: ${f.error}`)];

            if (allErrors.length > 0) {
                // eslint-disable-next-line no-console
                console.error(`CustomerMapping upload: ${allErrors.length} row(s) failed`, allErrors);
            }

            const summary = `${result.success} of ${dataRows.length} record(s) uploaded successfully.`;
            const notes: string[] = [];
            if (unrecognizedHeaders.length > 0) {
                notes.push(`Ignored unrecognized column(s): ${unrecognizedHeaders.join(", ")}.`);
            }
            if (allErrors.length > 0) {
                const maxShown = 10;
                const shown = allErrors.slice(0, maxShown).join(" | ");
                const remainder = allErrors.length - maxShown;
                notes.push(
                    `${allErrors.length} row(s) failed: ${shown}${remainder > 0 ? ` (and ${remainder} more — see browser console for the full list)` : ""}`
                );
            }

            this.setState({
                uploading: false,
                message: {
                    type: allErrors.length > 0 ? (result.success > 0 ? MessageBarType.warning : MessageBarType.error) : MessageBarType.success,
                    text: [summary, ...notes].join(" ")
                }
            });

            if (result.success > 0 && this.props.onUploadComplete) {
                this.props.onUploadComplete();
            }
        } catch (error) {
            this.setState({ uploading: false, message: { type: MessageBarType.error, text: error.message } });
        }
    };

    render(): React.ReactElement {
        const { uploading, message } = this.state;
        return (
            <Stack tokens={{ childrenGap: 8 }}>
                <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 10 }}>
                    <input
                        type="file"
                        ref={this.fileInputRef}
                        accept=".xlsx,.xls"
                        style={{ display: "none" }}
                        onChange={this.handleFileChange}
                    />
                    <PrimaryButton
                        text={uploading ? "Uploading..." : "Upload Excel"}
                        iconProps={{ iconName: "ExcelDocument" }}
                        onClick={this.triggerFileSelect}
                        disabled={uploading}
                    />
                    <DefaultButton
                        text="Download Template"
                        iconProps={{ iconName: "Download" }}
                        onClick={this.handleDownloadTemplate}
                        disabled={uploading}
                    />
                    {uploading && <Spinner size={SpinnerSize.small} label="Processing file..." />}
                </Stack>
                {message && (
                    <MessageBar
                        messageBarType={message.type}
                        onDismiss={() => this.setState({ message: null })}
                        isMultiline={true}
                        styles={{ root: { maxWidth: 700 } }}
                    >
                        <Text>{message.text}</Text>
                    </MessageBar>
                )}
            </Stack>
        );
    }
}

export default ExcelUploadComponent;
