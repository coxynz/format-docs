/**
 * Parser Module
 * Handles spreadsheet parsing for .xlsx and .csv files
 */

const Parser = {
    /**
     * Parse uploaded file and extract rows as objects
     * @param {File} file - The uploaded file
     * @returns {Promise<{headers: string[], rows: Object[]}>}
     */
    async parseFile(file) {
        const extension = file.name.split('.').pop().toLowerCase();

        if (extension === 'csv') {
            return this.parseCSV(file);
        } else if (extension === 'xlsx' || extension === 'xls') {
            return this.parseExcel(file);
        } else {
            throw new Error(`Unsupported file format: .${extension}`);
        }
    },

    /**
     * Parse Excel file using SheetJS
     * @param {File} file 
     * @returns {Promise<{headers: string[], rows: Object[]}>}
     */
    async parseExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Get first sheet
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];

                    // Convert to JSON with headers
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '', // Default empty cells to empty string
                        raw: false // Use formatted strings (fixes date format issue)
                    });

                    if (jsonData.length < 2) {
                        throw new Error('Spreadsheet must have at least a header row and one data row');
                    }

                    const headers = jsonData[0].map(h => String(h).trim());
                    const rows = [];

                    // Convert each row to an object using headers as keys
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        // Skip completely empty rows
                        if (row.every(cell => !cell || String(cell).trim() === '')) {
                            continue;
                        }

                        const rowObj = {};
                        headers.forEach((header, index) => {
                            rowObj[header] = row[index] !== undefined ? String(row[index]) : '';
                        });
                        rows.push(rowObj);
                    }

                    resolve({ headers, rows });
                } catch (error) {
                    reject(new Error(`Failed to parse Excel file: ${error.message}`));
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    },

    /**
     * Parse CSV file
     * @param {File} file 
     * @returns {Promise<{headers: string[], rows: Object[]}>}
     */
    async parseCSV(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const text = e.target.result;
                    const lines = this.parseCSVText(text);

                    if (lines.length < 2) {
                        throw new Error('CSV must have at least a header row and one data row');
                    }

                    const headers = lines[0].map(h => h.trim());
                    const rows = [];

                    for (let i = 1; i < lines.length; i++) {
                        const line = lines[i];
                        // Skip empty lines
                        if (line.every(cell => !cell || cell.trim() === '')) {
                            continue;
                        }

                        const rowObj = {};
                        headers.forEach((header, index) => {
                            rowObj[header] = line[index] !== undefined ? line[index] : '';
                        });
                        rows.push(rowObj);
                    }

                    resolve({ headers, rows });
                } catch (error) {
                    reject(new Error(`Failed to parse CSV file: ${error.message}`));
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsText(file);
        });
    },

    /**
     * Parse CSV text handling quoted values and commas within quotes
     * @param {string} text 
     * @returns {string[][]}
     */
    parseCSVText(text) {
        const lines = [];
        let currentLine = [];
        let currentValue = '';
        let insideQuotes = false;

        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            const nextChar = text[i + 1];

            if (char === '"') {
                if (insideQuotes && nextChar === '"') {
                    // Escaped quote
                    currentValue += '"';
                    i++;
                } else {
                    // Toggle quote mode
                    insideQuotes = !insideQuotes;
                }
            } else if (char === ',' && !insideQuotes) {
                currentLine.push(currentValue);
                currentValue = '';
            } else if ((char === '\n' || (char === '\r' && nextChar === '\n')) && !insideQuotes) {
                currentLine.push(currentValue);
                lines.push(currentLine);
                currentLine = [];
                currentValue = '';
                if (char === '\r') i++; // Skip \n in \r\n
            } else if (char !== '\r') {
                currentValue += char;
            }
        }

        // Don't forget the last value/line
        if (currentValue || currentLine.length > 0) {
            currentLine.push(currentValue);
            lines.push(currentLine);
        }

        return lines;
    }
};

// Export for use in other modules
window.Parser = Parser;
