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


    // Template HTML (loaded from file)
    template: null,

    // Logo Base64 (cached for DOCX embedding)
    logoBase64: null,

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

            // Also load logo as base64 for DOCX embedding
            await this.loadLogoAsBase64();

            return this.template;
        } catch (error) {
            throw new Error(`Template loading failed: ${error.message}`);
        }
    },

    /**
     * Load logo and convert to base64 for embedding in DOCX
     */
    async loadLogoAsBase64() {
        try {
            const logoPath = 'Images/VEGA-logo_with-slogan-removebg-preview.png';
            const response = await fetch(logoPath);
            if (!response.ok) return;

            const blob = await response.blob();
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    this.logoBase64 = reader.result;
                    resolve(this.logoBase64);
                };
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            console.warn('Failed to load logo for base64 embedding:', error);
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
    /**
     * Generate DOCX from filled HTML
     * @param {string} filledHtml 
     * @returns {Blob}
     */
    generateDocx(filledHtml) {
        // Extract styles from the original template
        const styles = this.extractStyles(this.template);

        // Wrap in proper document structure for html-docx-js
        const fullHtml = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                ${styles}
            </head>
            <body>
                ${this.prepareHtmlForDocx(filledHtml)}
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
     * Extract <style> block from HTML content
     * @param {string} html 
     * @returns {string} - The style block including <style> tags
     */
    extractStyles(html) {
        if (!html) return '';
        const match = html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
        return match ? match[0] : '';
    },

    /**
     * Prepare HTML for DOCX by extracting body and embedding base64 images
     * @param {string} html 
     * @returns {string}
     */
    prepareHtmlForDocx(html) {
        let content = this.extractBodyContent(html);

        // Replace logo path with base64 if available
        if (this.logoBase64) {
            content = content.replace(
                /src=["']Images\/VEGA-logo_with-slogan-removebg-preview\.png["']/g,
                `src="${this.logoBase64}"`
            );
        }

        return content;
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
     * Generate single document from a row
     * @param {Object} row 
     * @param {number} index 
     * @returns {Promise<{filename: string, blob: Blob}>}
     */
    async generateSingle(row, index) {
        const filledHtml = this.fillTemplate(row);
        const docxBlob = this.generateDocx(filledHtml);
        const filename = this.generateFilename(row, index);

        return { filename, blob: docxBlob };
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
