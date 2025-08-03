/**
 * File Handler for managing file uploads and processing
 */
class FileHandler {
    constructor() {
        this.files = [];
        this.fileContents = [];
    }

    /**
     * Add files to the handler
     * @param {FileList} fileList - Files from input or drop event
     * @returns {Array} - Array of added files
     */
    addFiles(fileList) {
        const newFiles = [];
        
        for (let i = 0; i < fileList.length; i++) {
            const file = fileList[i];
            
            // Check if file is PDF or text
            if (file.type === 'application/pdf' || file.type === 'text/plain') {
                this.files.push(file);
                newFiles.push(file);
            }
        }
        
        return newFiles;
    }

    /**
     * Remove a file by index
     * @param {number} index - Index of file to remove
     */
    removeFile(index) {
        if (index >= 0 && index < this.files.length) {
            this.files.splice(index, 1);
            this.fileContents.splice(index, 1);
        }
    }

    /**
     * Process all files and extract their contents
     * @returns {Promise<Array>} - Array of file contents as text
     */
    async processAllFiles() {
        this.fileContents = [];
        
        const promises = this.files.map(file => this.processFile(file));
        const results = await Promise.all(promises);
        
        this.fileContents = results;
        return results;
    }

    /**
     * Process a single file and extract its content
     * @param {File} file - File to process
     * @returns {Promise<string>} - File content as text
     */
    async processFile(file) {
        if (file.type === 'application/pdf') {
            return await this.processPdf(file);
        } else if (file.type === 'text/plain') {
            return await this.processText(file);
        }
        
        return '';
    }

    /**
     * Process a text file
     * @param {File} file - Text file
     * @returns {Promise<string>} - File content
     */
    async processText(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = event => {
                resolve(event.target.result);
            };
            
            reader.onerror = error => {
                console.error('Error reading text file:', error);
                reject(error);
            };
            
            reader.readAsText(file);
        });
    }

    /**
     * Process a PDF file using pdf.js
     * @param {File} file - PDF file
     * @returns {Promise<string>} - Extracted text content
     */
    async processPdf(file) {
        try {
            // Convert the file to an ArrayBuffer
            const arrayBuffer = await file.arrayBuffer();
            
            // Load the PDF document
            const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
            const pdf = await loadingTask.promise;
            
            // Extract text from all pages
            let textContent = '';
            
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const content = await page.getTextContent();
                const pageText = content.items.map(item => item.str).join(' ');
                textContent += pageText + ' ';
            }
            
            return textContent.trim();
        } catch (error) {
            console.error('Error processing PDF:', error);
            return `[Error processing PDF: ${file.name}]`;
        }
    }

    /**
     * Get all file contents as context
     * @returns {Array<string>} - Array of file contents
     */
    getFileContents() {
        return this.fileContents;
    }
    
    /**
     * Get all files
     * @returns {Array<File>} - Array of files
     */
    getFiles() {
        return this.files;
    }
}

// Create and export a singleton instance
const fileHandler = new FileHandler();
