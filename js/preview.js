/**
 * Preview Module
 * Handles document preview rendering
 */

const Preview = {
    // Reference to the preview iframe
    iframe: null,

    // Current data
    rows: [],
    currentIndex: 0,

    /**
     * Initialize the preview module
     * @param {HTMLIFrameElement} iframeElement 
     */
    init(iframeElement) {
        this.iframe = iframeElement;
    },

    /**
     * Set the data rows for preview
     * @param {Object[]} rows 
     */
    setRows(rows) {
        this.rows = rows;
        this.currentIndex = 0;
    },

    /**
     * Render the current row in the preview
     */
    render() {
        if (!this.iframe || this.rows.length === 0) {
            return;
        }

        const row = this.rows[this.currentIndex];
        const filledHtml = Generator.fillTemplate(row);

        // Write to iframe
        const doc = this.iframe.contentDocument || this.iframe.contentWindow.document;
        doc.open();
        doc.write(filledHtml);
        doc.close();
    },

    /**
     * Navigate to the next row
     * @returns {boolean} - Whether navigation was successful
     */
    next() {
        if (this.currentIndex < this.rows.length - 1) {
            this.currentIndex++;
            this.render();
            return true;
        }
        return false;
    },

    /**
     * Navigate to the previous row
     * @returns {boolean} - Whether navigation was successful
     */
    previous() {
        if (this.currentIndex > 0) {
            this.currentIndex--;
            this.render();
            return true;
        }
        return false;
    },

    /**
     * Jump to a specific row index
     * @param {number} index 
     */
    goTo(index) {
        if (index >= 0 && index < this.rows.length) {
            this.currentIndex = index;
            this.render();
        }
    },

    /**
     * Get current navigation state
     * @returns {{current: number, total: number, hasPrev: boolean, hasNext: boolean}}
     */
    getState() {
        return {
            current: this.currentIndex + 1,
            total: this.rows.length,
            hasPrev: this.currentIndex > 0,
            hasNext: this.currentIndex < this.rows.length - 1
        };
    },

    /**
     * Get the current row data
     * @returns {Object|null}
     */
    getCurrentRow() {
        return this.rows[this.currentIndex] || null;
    },

    /**
     * Clear the preview
     */
    clear() {
        this.rows = [];
        this.currentIndex = 0;

        if (this.iframe) {
            const doc = this.iframe.contentDocument || this.iframe.contentWindow.document;
            doc.open();
            doc.write('<html><body></body></html>');
            doc.close();
        }
    }
};

// Export for use in other modules
window.Preview = Preview;
