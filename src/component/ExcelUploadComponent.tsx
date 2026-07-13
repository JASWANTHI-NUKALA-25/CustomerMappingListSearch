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
import { columnsConfig } from "../constants/ColumnsConfig";

interface IExcelUploadProps {
    context: WebPartContext;
    listName: string;
    onUploadComplete?: () => void;
}

interface IExcelUploadState {
    uploading: boolean;
    message: { type: MessageBarType; text: string } | null;
}

const REQUIRED_HEADERS = ["Customer Name", "Customer ID"];

const HEADER_TO_KEY: Record<string, string> = columnsConfig.reduce((map, col) => {
    map[col.name.trim().toLowerCase()] = (col.fieldName as string) || (col.key as string);
    return map;
}, {} as Record<string, string>);

const EXPECTED_HEADERS = columnsConfig.map(col => col.name);

class ExcelUploadComponent extends React.Component<IExcelUploadProps, IExcelUploadState> {
    private searchService: SearchService;
    private fileInputRef = React.createRef<HTMLInputElement>();

    constructor(props: IExcelUploadProps) {
        super(props);
        this.searchService = new SearchService(props.context, props.listName);
        this.state = { uploading: false, message: null };
    }

    private triggerFileSelect = (): void => {
        if (this.fileInputRef.current) {
            this.fileInputRef.current.click();
        }
    };

    private handleDownloadTemplate = (): void => {
        const worksheet = XLSX.utils.aoa_to_sheet([EXPECTED_HEADERS]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "CustomerMapping");
        XLSX.writeFile(workbook, "CustomerMapping_Template.xlsx");
    };

    private handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
        const file = event.target.files && event.target.files[0];
        event.target.value = ""; // allow re-selecting the same file
        if (!file) return;

        this.setState({ uploading: true, message: null });

        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            if (!sheetName) throw new Error("The workbook does not contain any sheets.");

            const sheet = workbook.Sheets[sheetName];
            const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
            if (rows.length === 0) throw new Error("The uploaded file is empty.");

            const headerRow = (rows[0] || []).map((h: any) => (h || "").toString().trim());
            const missingHeaders = EXPECTED_HEADERS.filter(
                expected => !headerRow.some(h => h.toLowerCase() === expected.toLowerCase())
            );

            if (missingHeaders.length > 0) {
                throw new Error(
                    `The file's columns don't match the SharePoint list. Missing column(s): ${missingHeaders.join(", ")}. Use "Download Template" to get the correct format.`
                );
            }

            const unrecognizedHeaders = headerRow.filter(h => h && !HEADER_TO_KEY[h.toLowerCase()]);

            const dataRows = rows.slice(1).filter(row => row.some(cell => (cell ?? "").toString().trim() !== ""));
            if (dataRows.length === 0) throw new Error("No data rows found in the uploaded file.");

            const entries: { row: number; data: Record<string, string> }[] = [];
            const rowErrors: string[] = [];

            dataRows.forEach((row, idx) => {
                const excelRowNumber = idx + 2; // header is row 1
                const item: Record<string, string> = {};
                headerRow.forEach((header, colIdx) => {
                    const key = header ? HEADER_TO_KEY[header.toLowerCase()] : undefined;
                    if (key) {
                        item[key] = (row[colIdx] ?? "").toString().trim();
                    }
                });

                if (!item["Title"] || !item["field_1"]) {
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

            const summary = `${result.success} of ${dataRows.length} record(s) uploaded successfully.`;
            const notes: string[] = [];
            if (unrecognizedHeaders.length > 0) {
                notes.push(`Ignored unrecognized column(s): ${unrecognizedHeaders.join(", ")}.`);
            }
            if (allErrors.length > 0) {
                notes.push(`${allErrors.length} row(s) failed: ${allErrors.join(" | ")}`);
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
