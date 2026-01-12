/**
 * Generator Module
 * Handles template population and DOCX generation using docxtemplater
 */

const Generator = {
    // Column to placeholder mapping configuration
    // Keys match Excel columns, Values match [Tag] in docx template
    mappings: {
        'Client': 'INSERT_CLIENT_NAME',
        'Desired Completion Date': 'INSERT_DATE',
        'Room Details (Workload)': 'INSERT_ROOM_DETAILS',
        'Functional Requirements': 'INSERT_FUNCTIONAL_REQUIREMENTS',
        'Control System Requirements': 'INSERT_CONTROL_SYSTEM',
        'Preferred Brands or Technology Standards': 'INSERT_BRANDS_STANDARDS',
        'Existing Equipment Integration': 'INSERT_EXISTING_INTEGRATION',
        'Network Strategy': 'INSERT_CLIENT_NETWORK_OR_DEDICATED_AV',
        'Cabling & Infrastructure': 'INSERT_CABLING_DETAILS',
        'Budgetary Estimates': 'INSERT_BUDGET',
        'Site Constraints': 'INSERT_SITE_CONSTRAINTS'
    },

    // Templates
    htmlTemplate: null, // For preview only
    docxTemplate: null, // For generation (binary)

    /**
     * Load HTML template for Preview
     * @param {string} path 
     */
    async loadHtmlTemplate(path) {
        try {
            const response = await fetch(path);
            if (!response.ok) throw new Error(`Failed to load HTML template: ${response.status}`);
            this.htmlTemplate = await response.text();
        } catch (error) {
            console.error('HTML Template error:', error);
            throw error;
        }
    },

    /**
     * Load DOCX template for Generation
     * @param {string} path 
     */
    async loadDocxTemplate(path) {
        try {
            const response = await fetch(path);
            if (!response.ok) {
                if (response.status === 404) {
                    console.warn('Template.docx not found. Please create it.');
                    return false; // Signal that it's missing
                }
                throw new Error(`Failed to load DOCX template: ${response.status}`);
            }
            this.docxTemplate = await response.arrayBuffer();
            return true;
        } catch (error) {
            console.error('DOCX Template error:', error);
            throw error;
        }
    },

    /**
     * Fill HTML template for Preview (Simple string replacement)
     * @param {Object} rowData 
     * @returns {string} Filled HTML
     */
    fillHtmlTemplate(row) {
        if (!this.htmlTemplate) return '';

        let html = this.htmlTemplate;

        // Map data to HTML placeholders [Key]
        // Using the same mappings as DOCX for consistency
        for (const [col, tag] of Object.entries(this.mappings)) {
            const val = row[col] || '';
            const placeholder = `[${tag}]`;
            html = html.split(placeholder).join(this.escapeHtml(val));
        }
        return html;
    },

    /**
     * Generate Single DOCX using docxtemplater
     * @param {Object} row 
     * @param {number} index 
     */
    async generateSingle(row, index) {
        if (!this.docxTemplate) {
            throw new Error('Template.docx is missing used for generation. Please ensure Templates/template.docx exists.');
        }

        const zip = new PizZip(this.docxTemplate);

        let doc;
        try {
            doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                delimiters: { start: '[', end: ']' }
            });
        } catch (error) {
            throw new Error(`Docxtemplater init failed: ${error.message}`);
        }

        // Prepare data for docxtemplater
        const data = {};
        for (const [col, tag] of Object.entries(this.mappings)) {
            data[tag] = row[col] || ''; // Assign value to {{Tag}}
        }

        // Render
        try {
            doc.render(data);
        } catch (error) {
            console.error('Render error:', error);
            throw new Error('Failed to render document. Check template placeholders.');
        }

        // Generate blob
        const blob = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE',
        });

        const filename = this.generateFilename(row, index);
        return { filename, blob };
    },

    /**
     * Generate Filename
     */
    generateFilename(row, index) {
        const clientName = row['Client'] || `Document_${index + 1}`;
        const sanitized = clientName.replace(/[^a-zA-Z0-9\s\-_]/g, '').replace(/\s+/g, '_').substring(0, 50);
        return `${sanitized}_Specification.docx`;
    },

    /**
     * Generate All Documents
     */
    async generateAll(rows, progressCallback) {
        const documents = [];
        for (let i = 0; i < rows.length; i++) {
            documents.push(await this.generateSingle(rows[i], i));
            if (progressCallback) progressCallback(i + 1, rows.length);
            if (i % 10 === 0) await new Promise(r => setTimeout(r, 10));
        }
        return documents;
    },

    /**
     * Create Zip
     */
    async createZip(documents) {
        const zip = new JSZip();
        const usedNames = new Set();

        for (const doc of documents) {
            let filename = doc.filename;
            let counter = 1;
            while (usedNames.has(filename)) {
                filename = doc.filename.replace('.docx', `_${counter}.docx`);
                counter++;
            }
            usedNames.add(filename);
            zip.file(filename, doc.blob);
        }

        return await zip.generateAsync({ type: 'blob' });
    },

    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
};

window.Generator = Generator;
