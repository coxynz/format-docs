/**
 * Main Application Controller
 * Coordinates file upload, parsing, preview, and generation
 */

const App = {
    // State
    parsedData: null,
    templateLoaded: false,

    // DOM Elements
    elements: {},

    /**
     * Initialize the application
     */
    async init() {
        this.cacheElements();
        this.bindEvents();
        await this.loadTemplate();

        console.log('Format Docs initialized');
    },

    /**
     * Cache DOM element references
     */
    cacheElements() {
        this.elements = {
            dropZone: document.getElementById('dropZone'),
            fileInput: document.getElementById('fileInput'),
            uploadSection: document.getElementById('upload-section'),
            previewSection: document.getElementById('preview-section'),
            previewFrame: document.getElementById('previewFrame'),
            fileName: document.getElementById('fileName'),
            rowCount: document.getElementById('rowCount'),
            currentRow: document.getElementById('currentRow'),
            totalRows: document.getElementById('totalRows'),
            prevRow: document.getElementById('prevRow'),
            nextRow: document.getElementById('nextRow'),
            generateBtn: document.getElementById('generateBtn'),
            newFileBtn: document.getElementById('newFileBtn'),
            loadingOverlay: document.getElementById('loadingOverlay'),
            loadingText: document.getElementById('loadingText'),
            toastContainer: document.getElementById('toastContainer')
        };

        // Initialize Preview module with iframe
        Preview.init(this.elements.previewFrame);
    },

    /**
     * Bind event listeners
     */
    bindEvents() {
        const { dropZone, fileInput, prevRow, nextRow, generateBtn, newFileBtn } = this.elements;

        // File input change
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Drop zone click
        dropZone.addEventListener('click', () => fileInput.click());

        // Drag and drop
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drag-over');
        });

        dropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.processFile(files[0]);
            }
        });

        // Navigation buttons
        prevRow.addEventListener('click', () => this.navigatePrev());
        nextRow.addEventListener('click', () => this.navigateNext());

        // Action buttons
        generateBtn.addEventListener('click', () => this.generateDocuments());
        newFileBtn.addEventListener('click', () => this.reset());

        // Keyboard navigation
        document.addEventListener('keydown', (e) => {
            if (this.parsedData) {
                if (e.key === 'ArrowLeft') this.navigatePrev();
                if (e.key === 'ArrowRight') this.navigateNext();
            }
        });
    },

    /**
     * Load the HTML template
     */
    async loadTemplate() {
        try {
            await Generator.loadTemplate('Templates/format-docs.html');
            this.templateLoaded = true;
        } catch (error) {
            console.error('Failed to load template:', error);
            this.showToast('Failed to load template. Please refresh the page.', 'error');
        }
    },

    /**
     * Handle file input selection
     */
    handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            this.processFile(file);
        }
    },

    /**
     * Process uploaded file
     */
    async processFile(file) {
        // Validate file type
        const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
            'text/csv'];
        const extension = file.name.split('.').pop().toLowerCase();

        if (!['xlsx', 'xls', 'csv'].includes(extension)) {
            this.showToast('Please upload an Excel (.xlsx) or CSV file', 'error');
            return;
        }

        this.showLoading('Parsing spreadsheet...');

        try {
            const result = await Parser.parseFile(file);

            if (result.rows.length === 0) {
                throw new Error('No data rows found in the spreadsheet');
            }

            this.parsedData = result;

            // Update UI
            this.elements.fileName.textContent = file.name;
            this.elements.rowCount.textContent = `${result.rows.length} row${result.rows.length !== 1 ? 's' : ''}`;

            // Initialize preview
            Preview.setRows(result.rows);
            Preview.render();

            // Update navigation
            this.updateNavigation();

            // Show preview section
            this.elements.uploadSection.classList.add('hidden');
            this.elements.previewSection.classList.remove('hidden');

            this.hideLoading();
            this.showToast(`Loaded ${result.rows.length} row${result.rows.length !== 1 ? 's' : ''} successfully`, 'success');

        } catch (error) {
            this.hideLoading();
            console.error('Parse error:', error);
            this.showToast(error.message, 'error');
        }
    },

    /**
     * Navigate to previous row
     */
    navigatePrev() {
        if (Preview.previous()) {
            this.updateNavigation();
        }
    },

    /**
     * Navigate to next row
     */
    navigateNext() {
        if (Preview.next()) {
            this.updateNavigation();
        }
    },

    /**
     * Update navigation UI based on current state
     */
    updateNavigation() {
        const state = Preview.getState();

        this.elements.currentRow.textContent = state.current;
        this.elements.totalRows.textContent = state.total;
        this.elements.prevRow.disabled = !state.hasPrev;
        this.elements.nextRow.disabled = !state.hasNext;
    },

    /**
     * Generate and download DOCX documents
     */
    async generateDocuments() {
        if (!this.parsedData || !this.templateLoaded) {
            this.showToast('No data loaded or template missing', 'error');
            return;
        }

        const rows = this.parsedData.rows;

        this.showLoading(`Generating document${rows.length > 1 ? 's' : ''}...`);

        try {
            const documents = await Generator.generateAll(rows, (current, total) => {
                this.elements.loadingText.textContent = `Generating document ${current} of ${total}...`;
            });

            if (documents.length === 1) {
                // Single file - direct download
                saveAs(documents[0].blob, documents[0].filename);
                this.showToast('Document downloaded successfully!', 'success');
            } else {
                // Multiple files - create ZIP
                this.elements.loadingText.textContent = 'Creating ZIP archive...';
                const zipBlob = await Generator.createZip(documents);
                saveAs(zipBlob, 'Specification_Documents.zip');
                this.showToast(`${documents.length} documents downloaded as ZIP!`, 'success');
            }

            this.hideLoading();

        } catch (error) {
            this.hideLoading();
            console.error('Generation error:', error);
            this.showToast(`Generation failed: ${error.message}`, 'error');
        }
    },

    /**
     * Reset to initial state
     */
    reset() {
        this.parsedData = null;
        this.elements.fileInput.value = '';

        Preview.clear();

        // Show upload, hide preview
        this.elements.uploadSection.classList.remove('hidden');
        this.elements.previewSection.classList.add('hidden');
    },

    /**
     * Show loading overlay
     */
    showLoading(message = 'Processing...') {
        this.elements.loadingText.textContent = message;
        this.elements.loadingOverlay.classList.remove('hidden');
    },

    /**
     * Hide loading overlay
     */
    hideLoading() {
        this.elements.loadingOverlay.classList.add('hidden');
    },

    /**
     * Show toast notification
     */
    showToast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;

        this.elements.toastContainer.appendChild(toast);

        // Auto-remove after 4 seconds
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100%)';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => App.init());
