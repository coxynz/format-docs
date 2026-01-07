/**
 * Generator Module
 * Handles template population and DOCX generation
 */

const Generator = {
    // Column to placeholder mapping configuration
    // Easy to update when columns or placeholders change
    mappings: {
        'Client': '[INSERT_CLIENT_NAME]',
        'Desired Completion Date': '[INSERT_DATE]',
        'Room Details (Workload)': '[INSERT_ROOM_DETAILS]',
        'Functional Requirements': '[INSERT_FUNCTIONAL_REQUIREMENTS]',
        'Control System Requirements': '[INSERT_CONTROL_SYSTEM]',
        'Preferred Brands or Technology Standards': '[INSERT_BRANDS_STANDARDS]',
        'Existing Equipment Integration': '[INSERT_EXISTING_INTEGRATION]',
        'Network Strategy': '[INSERT_CLIENT_NETWORK_OR_DEDICATED_AV]',
        'Cabling & Infrastructure': '[INSERT_CABLING_DETAILS]',
        'Budgetary Estimates': '[INSERT_BUDGET]',
        'Site Constraints': '[INSERT_SITE_CONSTRAINTS]'
    },

    // Special mapping for file placeholders (combines multiple columns)
    fileMappings: {
        columns: ['Site Photos', 'Client Drawings'],
        placeholder: '[LIST_UPLOADED_FILES_OR_LINKS_HERE]'
    },

    // Template HTML (loaded from file)
    template: null,

    /**
     * Load template from file or URL
     * @param {string} templatePath 
     */
    async loadTemplate(templatePath) {
        try {
            const response = await fetch(templatePath);
            if (!response.ok) {
                throw new Error(`Failed to load template: ${response.status}`);
            }
            this.template = await response.text();
            return this.template;
        } catch (error) {
            throw new Error(`Template loading failed: ${error.message}`);
        }
    },

    /**
     * Set template directly from string
     * @param {string} templateHtml 
     */
    setTemplate(templateHtml) {
        this.template = templateHtml;
    },

    /**
     * Fill template with data from a single row
     * @param {Object} rowData - Object with column names as keys
     * @returns {string} - Filled HTML template
     */
    fillTemplate(rowData) {
        if (!this.template) {
            throw new Error('Template not loaded');
        }

        let filledHtml = this.template;

        // Replace standard mappings
        for (const [column, placeholder] of Object.entries(this.mappings)) {
            const value = rowData[column] || '';
            // Escape HTML entities in the value to prevent XSS
            const escapedValue = this.escapeHtml(value);
            filledHtml = filledHtml.split(placeholder).join(escapedValue || placeholder);
        }

        // Handle file placeholder (combine Site Photos and Client Drawings)
        const fileValues = this.fileMappings.columns
            .map(col => rowData[col])
            .filter(val => val && val.trim())
            .join('\n');

        const fileContent = fileValues || 'No files uploaded';
        filledHtml = filledHtml.split(this.fileMappings.placeholder).join(this.escapeHtml(fileContent));

        return filledHtml;
    },

    /**
     * Escape HTML entities to prevent XSS
     * @param {string} text 
     * @returns {string}
     */
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    /**
     * Generate DOCX from filled HTML
     * @param {string} filledHtml 
     * @returns {Blob}
     */
    generateDocx(filledHtml) {
        // Wrap in proper document structure for html-docx-js
        const fullHtml = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    /* Inline styles for DOCX compatibility */
                    body {
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        line-height: 1.6;
                        color: #333;
                        padding: 40px;
                    }
                    .container {
                        max-width: 900px;
                        margin: 0 auto;
                    }
                    header {
                        border-bottom: 3px solid #2c3e50;
                        margin-bottom: 30px;
                        padding-bottom: 10px;
                    }
                    h1 {
                        color: #2c3e50;
                        font-size: 24pt;
                        text-transform: uppercase;
                        letter-spacing: 1px;
                    }
                    .project-meta {
                        margin-bottom: 40px;
                        background: #f8f9fa;
                        padding: 20px;
                    }
                    .section {
                        margin-bottom: 30px;
                    }
                    .section-title {
                        font-size: 14pt;
                        color: #2c3e50;
                        border-bottom: 1px solid #dee2e6;
                        margin-bottom: 10px;
                        padding-bottom: 5px;
                        font-weight: bold;
                    }
                    .question-label {
                        font-weight: 600;
                        color: #34495e;
                        display: block;
                        margin-bottom: 5px;
                    }
                    .instruction-text {
                        font-size: 0.9em;
                        color: #666;
                        font-style: italic;
                        margin-bottom: 10px;
                    }
                    .response-box {
                        background-color: #fff;
                        padding: 10px 15px;
                        border-left: 4px solid #3498db;
                        margin-bottom: 20px;
                        min-height: 20px;
                    }
                    .file-placeholder {
                        border: 2px dashed #dee2e6;
                        padding: 20px;
                        text-align: center;
                        color: #666;
                    }
                    footer {
                        margin-top: 50px;
                        font-size: 0.8em;
                        color: #999;
                        text-align: center;
                    }
                </style>
            </head>
            <body>
                ${this.extractBodyContent(filledHtml)}
            </body>
            </html>
        `;

        // Use html-docx-js to convert
        const converted = htmlDocx.asBlob(fullHtml, {
            orientation: 'portrait',
            margins: {
                top: 720,    // 0.5 inch in twips
                right: 720,
                bottom: 720,
                left: 720
            }
        });

        return converted;
    },

    /**
     * Extract body content from full HTML document
     * @param {string} html 
     * @returns {string}
     */
    extractBodyContent(html) {
        const match = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
        return match ? match[1] : html;
    },

    /**
     * Generate filename for a row
     * @param {Object} rowData 
     * @param {number} index 
     * @returns {string}
     */
    generateFilename(rowData, index) {
        const clientName = rowData['Client'] || `Document_${index + 1}`;
        // Sanitize filename
        const sanitized = clientName
            .replace(/[^a-zA-Z0-9\s\-_]/g, '')
            .replace(/\s+/g, '_')
            .substring(0, 50);
        return `${sanitized}_Specification.docx`;
    },

    /**
     * Generate all documents from parsed rows
     * @param {Object[]} rows 
     * @param {Function} progressCallback - Called with (current, total)
     * @returns {Promise<{filename: string, blob: Blob}[]>}
     */
    async generateAll(rows, progressCallback) {
        const documents = [];

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const filledHtml = this.fillTemplate(row);
            const docxBlob = this.generateDocx(filledHtml);
            const filename = this.generateFilename(row, i);

            documents.push({ filename, blob: docxBlob });

            if (progressCallback) {
                progressCallback(i + 1, rows.length);
            }

            // Small delay to prevent UI freezing on large datasets
            if (i % 5 === 0) {
                await new Promise(r => setTimeout(r, 10));
            }
        }

        return documents;
    },

    /**
     * Bundle multiple documents into a ZIP file
     * @param {{filename: string, blob: Blob}[]} documents 
     * @returns {Promise<Blob>}
     */
    async createZip(documents) {
        const zip = new JSZip();

        // Track filenames to avoid duplicates
        const usedNames = new Set();

        for (const doc of documents) {
            let filename = doc.filename;
            let counter = 1;

            // Handle duplicate filenames
            while (usedNames.has(filename)) {
                const baseName = doc.filename.replace('.docx', '');
                filename = `${baseName}_${counter}.docx`;
                counter++;
            }

            usedNames.add(filename);
            zip.file(filename, doc.blob);
        }

        return await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
    }
};

// Export for use in other modules
window.Generator = Generator;
