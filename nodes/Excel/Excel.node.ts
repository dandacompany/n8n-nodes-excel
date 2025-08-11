import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
	ILoadOptionsFunctions,
	INodePropertyOptions,
	IDataObject,
	NodeConnectionType,
	INodeProperties,
	INodeParameters,
} from 'n8n-workflow';
import { DynamicTool } from '@langchain/core/tools';
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import * as csv from 'fast-csv';

// data 디렉토리 경로 설정
const dataDir = path.join(process.cwd(), 'data');

// data 디렉토리가 없으면 생성
if (!fs.existsSync(dataDir)) {
	fs.mkdirSync(dataDir, { recursive: true });
}

/**
 * Parses CSV data from a string into an array of objects.
 */
async function parseCsvData(csvData: string): Promise<any[]> {
	return new Promise((resolve, reject) => {
		const rows: any[] = [];
		csv.parseString(csvData, { headers: true })
			.on('error', error => reject(error))
			.on('data', row => rows.push(row))
			.on('end', () => resolve(rows));
	});
}

/**
 * Reads a CSV file and returns it as an XLSX.WorkBook object.
 */
async function readCsvAsWorkbook(filePath: string): Promise<XLSX.WorkBook> {
	const csvData = fs.readFileSync(filePath, 'utf-8');
	const rows = await parseCsvData(csvData);
	const ws = XLSX.utils.json_to_sheet(rows);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
	return wb;
}

/**
 * Gets the actual sheet name from a workbook. For CSVs, it's always the first sheet.
 */
function getSheetName(workbook: XLSX.WorkBook, requestedSheetName: string): string | undefined {
	if (workbook.SheetNames.length === 1) {
		return workbook.SheetNames[0];
	}
	return workbook.SheetNames.find(name => name === requestedSheetName);
}

/**
 * Writes a workbook to a file, handling both .xlsx and .csv formats.
 */
function writeWorkbook(workbook: XLSX.WorkBook, filePath: string, fileName: string): void {
	if (fileName.toLowerCase().endsWith('.csv')) {
		const sheetName = workbook.SheetNames[0];
		const sheet = workbook.Sheets[sheetName];
		const csvOutput = XLSX.utils.sheet_to_csv(sheet);
		fs.writeFileSync(filePath, csvOutput);
	} else {
		XLSX.writeFile(workbook, filePath);
	}
}

/**
 * Apply filters to data based on filter conditions
 */
function applyFilters(data: IDataObject[], filters: { conditions?: Array<{ column: string; operator: string; value: string; }> }): IDataObject[] {
	if (!filters.conditions || filters.conditions.length === 0) {
		return data;
	}
	
	return data.filter(row => {
		return filters.conditions!.every(filter => {
			const cellValue = String(row[filter.column] || '');
			const filterValue = String(filter.value || '');
			
			switch (filter.operator) {
				case 'contains':
					return cellValue.toLowerCase().includes(filterValue.toLowerCase());
				case 'notContains':
					return !cellValue.toLowerCase().includes(filterValue.toLowerCase());
				case 'equals':
					return cellValue === filterValue;
				case 'notEquals':
					return cellValue !== filterValue;
				case 'startsWith':
					return cellValue.toLowerCase().startsWith(filterValue.toLowerCase());
				case 'endsWith':
					return cellValue.toLowerCase().endsWith(filterValue.toLowerCase());
				case 'isEmpty':
					return cellValue === '' || cellValue === null || cellValue === undefined;
				case 'isNotEmpty':
					return cellValue !== '' && cellValue !== null && cellValue !== undefined;
				case 'greaterThan':
					return parseFloat(cellValue) > parseFloat(filterValue);
				case 'lessThan':
					return parseFloat(cellValue) < parseFloat(filterValue);
				case 'greaterOrEqual':
					return parseFloat(cellValue) >= parseFloat(filterValue);
				case 'lessOrEqual':
					return parseFloat(cellValue) <= parseFloat(filterValue);
				default:
					return true;
			}
		});
	});
}

export class Excel implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel/CSV',
		name: 'excel',
		icon: 'file:spreadsheet.png',
		group: ['input', 'tools'],
		version: 1,
		description: 'Operate on local Excel/CSV files like a database. Compatible with AI Agent nodes as a tool.',
		defaults: {
			name: 'Excel/CSV',
		},
		inputs: ['main'] as NodeConnectionType[],
		outputs: ['main'] as NodeConnectionType[],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [ 
					{ name: 'File', value: 'file' }, 
					{ name: 'Data', value: 'data' }
				],
				default: 'data',
			},
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['file'] } },
				options: [
					{ name: 'Create', value: 'create' },
					{ name: 'Delete', value: 'delete' },
					{ name: 'Download', value: 'download' },
					{ name: 'Upload', value: 'upload' },
				],
				default: 'create',
			},
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['data'] } },
				options: [
					{ name: 'Read', value: 'read' },
					{ name: 'Add Row', value: 'addRow' },
					{ name: 'Update Rows', value: 'updateRows' },
					{ name: 'Delete Rows', value: 'deleteRows' },
					{ name: 'Clear Data', value: 'clearData' },
				],
				default: 'read',
			},
			// File Name for Data operations (with AI support)
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'options',
				description: 'Choose or specify the file to operate on.',
				required: true,
				typeOptions: {
					loadOptionsMethod: 'getFiles',
					allowCustomValue: true,
				},
				default: '',
				displayOptions: {
					show: {
						resource: ['data'],
					},
				},
			},
			// File Name for File Delete/Download operations
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'options',
				description: 'Choose the file to operate on.',
				required: true,
				typeOptions: {
					loadOptionsMethod: 'getFiles',
				},
				default: '',
				displayOptions: {
					show: {
						resource: ['file'],
						operation: ['delete', 'download'],
					},
				},
			},
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'string',
				required: true,
				default: 'new_file',
				description: 'Name of the new file to create, without the extension.',
				displayOptions: { show: { resource: ['file'], operation: ['create'] } },
			},
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'string',
				required: true,
				default: 'uploaded_file.xlsx',
				description: 'Name for the uploaded file. Must include .xlsx or .csv extension.',
				displayOptions: { show: { resource: ['file'], operation: ['upload'] } },
			},
			{
				displayName: 'Upload Source',
				name: 'uploadSource',
				type: 'options',
				options: [
					{ name: 'From Previous Node', value: 'previousNode' },
					{ name: 'From File Path', value: 'filePath' },
				],
				default: 'previousNode',
				description: 'Choose whether to upload from a previous node\'s binary data or from a local file path.',
				displayOptions: { show: { resource: ['file'], operation: ['upload'] } },
			},
			{
				displayName: 'Source File Path',
				name: 'sourceFilePath',
				type: 'string',
				default: '',
				required: true,
				description: 'The full path to the file to upload.',
				displayOptions: { show: { resource: ['file'], operation: ['upload'], uploadSource: ['filePath'] } },
			},
			{
				displayName: 'File Type',
				name: 'fileType',
				type: 'options',
				options: [ { name: 'Excel (.xlsx)', value: 'xlsx' }, { name: 'CSV (.csv)', value: 'csv' } ],
				default: 'xlsx',
				displayOptions: { show: { resource: ['file'], operation: ['create'] } },
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'options',
				default: 'Sheet1',
				description: 'Sheet to operate on. Required for .xlsx files. Can be specified manually for AI agents.',
				displayOptions: {
					show: {
						resource: [
							'data',
						],
					},
					hide: {
						fileName: [
							'.*\\.[cC][sS][vV]$',
						]
					}
				},
				typeOptions: {
					loadOptionsMethod: 'getSheetNames',
					loadOptionsDependsOn: ['fileName'],
					allowCustomValue: true,
				},
			},
			{
				displayName: 'Limit',
				name: 'limit',
				type: 'number',
				typeOptions: { minValue: 0 },
				default: 10,
				description: 'Number of rows to return. Use 0 for all rows.',
				displayOptions: { show: { operation: ['read'] } },
			},
			// Filter options for Read operation
			{
				displayName: 'Filters',
				name: 'filters',
				type: 'fixedCollection',
				placeholder: 'Add Filter',
				default: {},
				description: 'Filter rows based on column values',
				displayOptions: {
					show: {
						resource: ['data'],
						operation: ['read'],
					},
				},
				typeOptions: {
					multipleValues: true,
				},
				options: [
					{
						displayName: 'Filter',
						name: 'conditions',
						values: [
							{
								displayName: 'Column',
								name: 'column',
								type: 'options',
								default: '',
								description: 'Column name to filter by',
								required: true,
								typeOptions: {
									loadOptionsMethod: 'getColumns',
									loadOptionsDependsOn: ['fileName', 'sheetName'],
									allowCustomValue: true,
								},
							},
							{
								displayName: 'Operator',
								name: 'operator',
								type: 'options',
								options: [
									{ name: 'Contains', value: 'contains' },
									{ name: 'Not Contains', value: 'notContains' },
									{ name: 'Equals', value: 'equals' },
									{ name: 'Not Equals', value: 'notEquals' },
									{ name: 'Starts With', value: 'startsWith' },
									{ name: 'Ends With', value: 'endsWith' },
									{ name: 'Is Empty', value: 'isEmpty' },
									{ name: 'Is Not Empty', value: 'isNotEmpty' },
									{ name: 'Greater Than', value: 'greaterThan' },
									{ name: 'Less Than', value: 'lessThan' },
									{ name: 'Greater or Equal', value: 'greaterOrEqual' },
									{ name: 'Less or Equal', value: 'lessOrEqual' },
								],
								default: 'contains',
								description: 'Filter operator',
							},
							{
								displayName: 'Value',
								name: 'value',
								type: 'string',
								default: '',
								description: 'Value to compare (not needed for isEmpty/isNotEmpty)',
								displayOptions: {
									hide: {
										operator: ['isEmpty', 'isNotEmpty'],
									},
								},
							},
						],
					},
				],
			},
			{
				displayName: 'Sort',
				name: 'sort',
				type: 'fixedCollection',
				placeholder: 'Add Sort',
				default: {},
				description: 'Sort rows by column values',
				displayOptions: {
					show: {
						resource: ['data'],
						operation: ['read'],
					},
				},
				typeOptions: {
					multipleValues: false,
				},
				options: [
					{
						displayName: 'Sort',
						name: 'fields',
						values: [
							{
								displayName: 'Column',
								name: 'column',
								type: 'options',
								default: '',
								description: 'Column to sort by',
								typeOptions: {
									loadOptionsMethod: 'getColumns',
									loadOptionsDependsOn: ['fileName', 'sheetName'],
									allowCustomValue: true,
								},
							},
							{
								displayName: 'Direction',
								name: 'direction',
								type: 'options',
								options: [
									{ name: 'Ascending', value: 'asc' },
									{ name: 'Descending', value: 'desc' },
								],
								default: 'asc',
							},
						],
					},
				],
			},
			{
				displayName: 'Key Column',
				name: 'keyColumn',
				type: 'string',
				default: 'ID',
				displayOptions: { show: { operation: ['updateRow'] } },
			},
			{
				displayName: 'Key Value',
				name: 'keyValue',
				type: 'string',
				default: '',
				displayOptions: { show: { operation: ['updateRow'] } },
			},
			{
				displayName: 'Row Data',
				name: 'rowData',
				type: 'fixedCollection',
				placeholder: 'Add Column Data',
				default: {},
				description: 'Data for the row. Columns are loaded from the selected file and sheet.',
				displayOptions: {
					show: {
						resource: ['data'],
						operation: ['addRow', 'updateRows'],
						fileName: [{ _cnd: { exists: true } }],
						sheetName: [{ _cnd: { exists: true } }],
					},
				},
				typeOptions: {
					multipleValues: true,
				},
				options: [
					{
						displayName: 'Column Data',
						name: 'entries',
						values: [
							{
								displayName: 'Column',
								name: 'column',
								type: 'options',
								default: '',
								typeOptions: {
									loadOptionsMethod: 'getFieldsAsOptions',
									loadOptionsDependsOn: ['fileName', 'sheetName'],
								},
								description: 'Name of the column.',
							},
							{
								displayName: 'Value',
								name: 'value',
								type: 'string',
								default: '',
								description: 'Value for the column.',
							},
						],
					},
				],
			},
			{
				displayName: 'Columns',
				name: 'fileColumns',
				type: 'string',
				default: [],
				placeholder: 'Column Name',
				description: 'Define the columns for the new file. If empty, an empty file will be created.',
				displayOptions: {
					show: {
						resource: ['file'],
						operation: ['create'],
					},
				},
				typeOptions: {
					multipleValues: true,
					multipleValueButtonText: 'Add Column',
				},
			},
		],
	};

	async update(this: ILoadOptionsFunctions, changedParams: INodeParameters): Promise<INodeParameters> {
		const resource = this.getCurrentNodeParameter('resource') as string;
		const operation = this.getCurrentNodeParameter('operation') as string;

		if (resource === 'data' && (operation === 'addRow' || operation === 'updateRow')) {
			if (changedParams.fileName || changedParams.sheetName) {
				const fileName = this.getCurrentNodeParameter('fileName') as string;
				const sheetNameParam = this.getCurrentNodeParameter('sheetName') as string;

				if (fileName) {
					const filePath = path.join(dataDir, fileName);
					if (fs.existsSync(filePath)) {
						try {
							const workbook = fileName.toLowerCase().endsWith('.csv') ? await readCsvAsWorkbook(filePath) : XLSX.readFile(filePath);
							const sheetName = getSheetName(workbook, sheetNameParam ?? '');

							if (sheetName) {
								const sheet = workbook.Sheets[sheetName];
								const headerData = ((XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] as string[]) || []).filter(h => h);

								if (headerData.length > 0) {
									const currentParams = this.getCurrentNodeParameters() as INodeParameters;
									const rowData = currentParams.rowData as { entries: { column: string }[] } | undefined;
									const existingColumns = rowData?.entries?.map(e => e.column) ?? [];

									const headersChanged = headerData.length !== existingColumns.length || headerData.some(h => !existingColumns.includes(h));

									if (changedParams.fileName || headersChanged || existingColumns.length === 0) {
										return {
											rowData: {
												entries: headerData.map(h => ({ column: h, value: '' })),
											},
										};
									}
								}
							}
						} catch (e) {
							// Silently fail, user might be typing filename
						}
					}
				}
			}
		}
		return changedParams;
	}

	methods = {
		loadOptions: {
			async getFiles(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const files = fs.readdirSync(dataDir).filter(f => f.toLowerCase().endsWith('.xlsx') || f.toLowerCase().endsWith('.csv'));
				return files.map(file => ({ name: file, value: file }));
			},
			async getSheetNames(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const fileName = this.getCurrentNodeParameter('fileName') as string;
				if (!fileName || !fileName.toLowerCase().endsWith('.xlsx')) return [];
				const filePath = path.join(dataDir, fileName);
				if (!fs.existsSync(filePath)) return [];
				try {
					const workbook = XLSX.readFile(filePath);
					return workbook.SheetNames.map(name => ({ name: name, value: name }));
				} catch (e) {
					return [];
				}
			},
			async getFieldsAsOptions(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const fileName = this.getCurrentNodeParameter('fileName') as string;
				if (!fileName) return [];

				const filePath = path.join(dataDir, fileName);
				if (!fs.existsSync(filePath)) return [];

				try {
					const workbook = fileName.toLowerCase().endsWith('.csv') ? await readCsvAsWorkbook(filePath) : XLSX.readFile(filePath);
					let sheetName = getSheetName(workbook, this.getCurrentNodeParameter('sheetName') as string ?? '');

					if (!sheetName && workbook.SheetNames.length > 0) {
						sheetName = workbook.SheetNames[0];
					}

					if (!sheetName) return [];

					const sheet = workbook.Sheets[sheetName];
					const headerData = (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] as string[]) || [];
					if (!headerData) return [];

					return headerData.filter(h => h).map(h => ({ name: h, value: h }));
				} catch (e) {
					return [];
				}
			},
			async getColumns(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const fileName = this.getCurrentNodeParameter('fileName') as string;
				if (!fileName) return [];

				const filePath = path.join(dataDir, fileName);
				if (!fs.existsSync(filePath)) return [];

				try {
					const workbook = fileName.toLowerCase().endsWith('.csv') ? await readCsvAsWorkbook(filePath) : XLSX.readFile(filePath);
					let sheetName = fileName.toLowerCase().endsWith('.csv') 
						? 'Sheet1' 
						: (this.getCurrentNodeParameter('sheetName') as string || workbook.SheetNames[0]);

					if (!sheetName || !workbook.Sheets[sheetName]) return [];

					const sheet = workbook.Sheets[sheetName];
					const headerData = (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] as string[]) || [];
					
					return headerData.filter(h => h).map(h => ({ name: h, value: h }));
				} catch (e) {
					return [];
				}
			},
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			try {
				const resource = this.getNodeParameter('resource', i) as string;

				// Handle AI Tool mode
				if (resource === 'tool') {
					const toolName = this.getNodeParameter('toolName', i) as string;
					const toolDescription = this.getNodeParameter('toolDescription', i) as string;
					const availableFiles = this.getNodeParameter('availableFiles', i, '') as string;

					// Build tool description with available files and operations
					const files = availableFiles ? availableFiles.split(',').map(f => f.trim()) : fs.readdirSync(dataDir).filter(f => 
						f.toLowerCase().endsWith('.xlsx') || f.toLowerCase().endsWith('.csv')
					);
					let description = toolDescription;
					if (files.length > 0) {
						description += `\n\nAvailable files: ${files.join(', ')}`;
					}
					description += `\n\nSupported operations:
- read_data: Read data from a file (specify filename and optionally limit)
- add_row: Add a new row to a file (specify filename, sheet name for xlsx, and data)
- update_row: Update an existing row (specify filename, sheet name, key column, key value, and new data)
- create_file: Create a new file (specify filename, file type, and optional columns)
- list_files: List all available Excel/CSV files
- get_file_info: Get information about a specific file (sheets, columns, etc.)

Input format: JSON string with 'operation', 'filename', and other relevant parameters.
Example: {"operation": "read_data", "filename": "data.xlsx", "sheet": "Sheet1", "limit": 10}`;

					// Create the DynamicTool for AI agents
					const excelTool = new DynamicTool({
						name: toolName,
						description: description,
						func: async (input: string) => {
							try {
								const params = JSON.parse(input);
								const operation = params.operation;

								switch (operation) {
									case 'list_files':
										const fileList = fs.readdirSync(dataDir).filter(f => 
											f.toLowerCase().endsWith('.xlsx') || f.toLowerCase().endsWith('.csv')
										);
										return fileList.length === 0 
											? 'No Excel or CSV files found in the data directory.'
											: `Available files:\n${fileList.map(f => `- ${f}`).join('\n')}`;
									
									case 'read_data':
										const { filename, sheet, limit = 10 } = params;
										if (!filename) return 'Error: filename parameter is required for read_data operation.';
										
										const filePath = path.join(dataDir, filename);
										if (!fs.existsSync(filePath)) return `Error: File '${filename}' not found.`;
										
										const workbook = filename.toLowerCase().endsWith('.csv') 
											? await readCsvAsWorkbook(filePath) 
											: XLSX.readFile(filePath);
										
										const sheetName = getSheetName(workbook, sheet || 'Sheet1');
										if (!sheetName) return `Error: Sheet '${sheet}' not found. Available sheets: ${workbook.SheetNames.join(', ')}`;

										const worksheet = workbook.Sheets[sheetName];
										const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' }) as IDataObject[];
										const limitedData = data.slice(0, Math.min(limit, 100));

										return `Data from '${filename}' (${sheet || sheetName}):\n${JSON.stringify(limitedData, null, 2)}\n\nTotal rows: ${data.length}, Showing: ${limitedData.length}`;
									
									default:
										return `Error: Unknown operation '${operation}'. Supported operations: list_files, read_data, add_row, update_row, create_file, get_file_info`;
								}
							} catch (error) {
								if (error instanceof SyntaxError) {
									return `Error: Invalid JSON input. Please provide a valid JSON string with operation and parameters.`;
								}
								return `Error: ${error instanceof Error ? error.message : String(error)}`;
							}
						},
					});

					returnData.push({
						json: { tool: excelTool },
						pairedItem: { item: i },
					});

					continue;
				}

				const operation = this.getNodeParameter('operation', i) as string;

				if (resource === 'file') {
					if (operation === 'create') {
						const fileType = this.getNodeParameter('fileType', i) as string;
						const fileName = `${this.getNodeParameter('fileName', i)}.${fileType}`;

						const headers = this.getNodeParameter('fileColumns', i, []) as string[];

						const ws = XLSX.utils.json_to_sheet([], { header: headers.filter(Boolean) });
						const wb = XLSX.utils.book_new();
						XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
						writeWorkbook(wb, path.join(dataDir, fileName), fileName);
						returnData.push({ json: { success: true, message: `File '${fileName}' created.` } });
					} else if (operation === 'delete') {
						const fileName = this.getNodeParameter('fileName', i) as string;
						const filePath = path.join(dataDir, fileName);
						if (fs.existsSync(filePath)) {
							fs.unlinkSync(filePath);
							returnData.push({ json: { success: true, message: `File '${fileName}' deleted.` } });
						} else {
							throw new NodeOperationError(this.getNode(), `File '${fileName}' not found.`);
						}
					} else if (operation === 'upload') {
						const fileName = this.getNodeParameter('fileName', i) as string;
						if (!fileName || (!fileName.toLowerCase().endsWith('.xlsx') && !fileName.toLowerCase().endsWith('.csv'))) {
							throw new NodeOperationError(this.getNode(), 'File name is required and must end with .xlsx or .csv');
						}

						const uploadSource = this.getNodeParameter('uploadSource', i, 'previousNode') as string;
						let fileBuffer: Buffer;

						if (uploadSource === 'filePath') {
							const sourceFilePath = this.getNodeParameter('sourceFilePath', i) as string;
							if (!sourceFilePath) {
								throw new NodeOperationError(this.getNode(), 'Source File Path is required.');
							}
							if (!fs.existsSync(sourceFilePath)) {
								throw new NodeOperationError(this.getNode(), `Source file not found at: ${sourceFilePath}`);
							}
							fileBuffer = fs.readFileSync(sourceFilePath);
						} else { // 'previousNode'
							if (items[i].binary === undefined) {
								throw new NodeOperationError(this.getNode(), 'No binary data found on the input to upload. Please connect a node that returns a file.');
							}

							// 'data' is the default property name for binary data in n8n
							fileBuffer = await this.helpers.getBinaryDataBuffer(i, 'data');
						}

						const filePath = path.join(dataDir, fileName);
						fs.writeFileSync(filePath, fileBuffer);
						returnData.push({ json: { success: true, message: `File '${fileName}' uploaded successfully.` } });
					} else if (operation === 'download') {
						const fileName = this.getNodeParameter('fileName', i) as string;
						if (!fileName) {
							throw new NodeOperationError(this.getNode(), 'File name is not specified.');
						}
						const filePath = path.join(dataDir, fileName);
						if (!fs.existsSync(filePath)) {
							throw new NodeOperationError(this.getNode(), `File '${fileName}' not found.`);
						}

						const fileData = fs.readFileSync(filePath);
						const binaryData = await this.helpers.prepareBinaryData(fileData, fileName);

						returnData.push({
							json: {},
							binary: {
								data: binaryData,
							},
						});
					}
				} else if (resource === 'data') {
					const fileName = this.getNodeParameter('fileName', i) as string;
					const filePath = path.join(dataDir, fileName);
					if (!fs.existsSync(filePath)) throw new NodeOperationError(this.getNode(), `File '${fileName}' not found.`);

					const workbook = fileName.toLowerCase().endsWith('.csv') ? await readCsvAsWorkbook(filePath) : XLSX.readFile(filePath);
					const requestedSheet = this.getNodeParameter('sheetName', i, 'Sheet1') as string;
					const currentSheetName = getSheetName(workbook, requestedSheet);
					if (!currentSheetName) throw new NodeOperationError(this.getNode(), `Sheet '${requestedSheet}' not found.`);
					const sheet = workbook.Sheets[currentSheetName];

					if (operation === 'read') {
						const limit = this.getNodeParameter('limit', i, 10) as number;
						let data = XLSX.utils.sheet_to_json(sheet, { defval: '' }) as IDataObject[];
						
						// Apply filters
						const filters = this.getNodeParameter('filters', i, {}) as {
							conditions?: Array<{
								column: string;
								operator: string;
								value: string;
							}>;
						};
						
						data = applyFilters(data, filters);
						
						// Apply sorting
						const sort = this.getNodeParameter('sort', i, {}) as {
							fields?: {
								column: string;
								direction: 'asc' | 'desc';
							};
						};
						
						if (sort.fields && sort.fields.column) {
							data.sort((a, b) => {
								const aVal = String(a[sort.fields!.column] || '');
								const bVal = String(b[sort.fields!.column] || '');
								
								// Try to compare as numbers first
								const aNum = parseFloat(aVal);
								const bNum = parseFloat(bVal);
								
								let comparison = 0;
								if (!isNaN(aNum) && !isNaN(bNum)) {
									comparison = aNum - bNum;
								} else {
									comparison = aVal.localeCompare(bVal);
								}
								
								return sort.fields!.direction === 'desc' ? -comparison : comparison;
							});
						}
						
						// Apply limit (0 means all rows)
						const finalData = limit === 0 ? data : data.slice(0, limit);
						returnData.push(...this.helpers.returnJsonArray(finalData));
					} else if (operation === 'addRow' || operation === 'updateRows' || operation === 'deleteRows') {
						let finalRowData: IDataObject = {};
						let newHeaders: string[] = [];
						
						// Only get rowData for addRow and updateRows operations
						if (operation === 'addRow' || operation === 'updateRows') {
							const rowDataCollection = this.getNodeParameter('rowData', i) as { entries: Array<{ column: string; value: any; }> };
							if (rowDataCollection?.entries) {
								for (const entry of rowDataCollection.entries) {
									if (entry.column) {
										finalRowData[entry.column] = entry.value;
									}
								}
							}
						}

						let data = XLSX.utils.sheet_to_json(sheet) as IDataObject[];
						const headerData = (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] as string[]) || [];
						newHeaders = [...headerData];

						if (operation === 'addRow' || operation === 'updateRows') {
							Object.keys(finalRowData).forEach(key => {
								if (!newHeaders.includes(key)) {
									newHeaders.push(key);
								}
							});
						}

						if (operation === 'addRow') {
							data.push(finalRowData);
							const newSheet = XLSX.utils.json_to_sheet(data, { header: newHeaders });
							workbook.Sheets[currentSheetName] = newSheet;
							writeWorkbook(workbook, filePath, fileName);
							returnData.push({ json: { success: true, message: 'Row added.' } });
						} else if (operation === 'updateRows') {
							// Get filter criteria
							const updateFilters = this.getNodeParameter('updateFilters', i, {}) as {
								conditions?: Array<{
									column: string;
									operator: string;
									value: string;
								}>;
							};
							
							// Find rows to update
							const rowsToUpdate = applyFilters(data, updateFilters);
							
							if (rowsToUpdate.length === 0) {
								throw new NodeOperationError(this.getNode(), 'No rows found matching the filter criteria.');
							}
							
							// Update matching rows
							let updatedCount = 0;
							data = data.map(row => {
								const isMatch = rowsToUpdate.some(matchRow => {
									// Check if this is the same row by comparing all values
									return Object.keys(matchRow).every(key => row[key] === matchRow[key]);
								});
								
								if (isMatch) {
									updatedCount++;
									return { ...row, ...finalRowData };
								}
								return row;
							});
							
							const newSheet = XLSX.utils.json_to_sheet(data, { header: newHeaders });
							workbook.Sheets[currentSheetName] = newSheet;
							writeWorkbook(workbook, filePath, fileName);
							returnData.push({ json: { success: true, message: `${updatedCount} row(s) updated.` } });
						} else if (operation === 'deleteRows') {
							// Get filter criteria
							const deleteFilters = this.getNodeParameter('updateFilters', i, {}) as {
								conditions?: Array<{
									column: string;
									operator: string;
									value: string;
								}>;
							};
							
							// Find rows to delete
							const rowsToDelete = applyFilters(data, deleteFilters);
							
							if (rowsToDelete.length === 0) {
								throw new NodeOperationError(this.getNode(), 'No rows found matching the filter criteria.');
							}
							
							// Filter out rows that match deletion criteria
							const deletedCount = rowsToDelete.length;
							data = data.filter(row => {
								const shouldDelete = rowsToDelete.some(deleteRow => {
									// Check if this is the same row by comparing all values
									return Object.keys(deleteRow).every(key => row[key] === deleteRow[key]);
								});
								return !shouldDelete;
							});
							
							// Keep headers even if all data is deleted
							const headers = (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] as string[]) || [];
							const newSheet = XLSX.utils.json_to_sheet(data, { header: headers });
							workbook.Sheets[currentSheetName] = newSheet;
							writeWorkbook(workbook, filePath, fileName);
							returnData.push({ json: { success: true, message: `${deletedCount} row(s) deleted.` } });
						}
					} else if (operation === 'clearData') {
						const headers = (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] as string[]) || [];
						workbook.Sheets[currentSheetName] = XLSX.utils.json_to_sheet([], { header: headers });
						writeWorkbook(workbook, filePath, fileName);
						returnData.push({ json: { success: true, message: 'Data cleared.' } });
					}
				}
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({ json: { error: error instanceof Error ? error.message : String(error) }, pairedItem: { item: i } });
					continue;
				}
				throw error;
			}
		}
		return [returnData];
	}

}