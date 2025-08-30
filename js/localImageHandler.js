/**
 * Local Image Handler for managing local folder images
 */
class LocalImageHandler {
    constructor() {
        this.selectedImages = new Map(); // Map of filename -> base64 data
        this.folderPath = '';
        this.imageCount = 0;
        this.supportedFormats = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'];
    }
    
    /**
     * Clear all selected images
     */
    clear() {
        this.selectedImages.clear();
        this.folderPath = '';
        this.imageCount = 0;
    }
    
    /**
     * Process files from folder selection
     * @param {FileList} files - Files from directory input
     * @returns {Promise<object>} - Result with image count and folder info
     */
    async processFolder(files) {
        this.clear();
        
        // Filter for image files only
        const imageFiles = Array.from(files).filter(file => {
            const extension = file.name.split('.').pop().toLowerCase();
            return this.supportedFormats.includes(extension);
        });
        
        if (imageFiles.length === 0) {
            return {
                success: false,
                message: 'No supported image files found in the selected folder',
                imageCount: 0
            };
        }
        
        // Extract folder path from first file
        if (imageFiles[0].webkitRelativePath) {
            const pathParts = imageFiles[0].webkitRelativePath.split('/');
            this.folderPath = pathParts[0];
        }
        
        // Process each image file
        const processPromises = imageFiles.map(file => this.processImageFile(file));
        await Promise.all(processPromises);
        
        this.imageCount = this.selectedImages.size;
        
        return {
            success: true,
            folderPath: this.folderPath,
            imageCount: this.imageCount,
            imageNames: Array.from(this.selectedImages.keys())
        };
    }
    
    /**
     * Process a single image file
     * @param {File} file - Image file to process
     * @returns {Promise<void>}
     */
    async processImageFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                // Extract just the filename without path
                const filename = file.name;
                // Store as base64 data URL
                this.selectedImages.set(filename, e.target.result);
                resolve();
            };
            
            reader.onerror = (error) => {
                console.error(`Error reading file ${file.name}:`, error);
                resolve(); // Continue processing other files
            };
            
            reader.readAsDataURL(file);
        });
    }
    
    /**
     * Match an image to a slide based on various strategies
     * @param {object} slideData - Slide data containing title and content
     * @param {number} slideIndex - Zero-based index of the slide (excluding title slide)
     * @returns {object|null} - Matched image data or null
     */
    matchImageForSlide(slideData, slideIndex) {
        if (this.selectedImages.size === 0) {
            return null;
        }
        
        const slideTitle = slideData.title || '';
        console.log(`Matching image for slide ${slideIndex}: "${slideTitle}"`);
        
        // Strategy 1: Exact title match (case-insensitive)
        const exactMatch = this.findExactTitleMatch(slideTitle);
        if (exactMatch) {
            console.log(`  Found exact match: ${exactMatch.filename}`);
            return exactMatch;
        }
        
        // Strategy 2: Slide number match (slide1, slide_1, slide-1, etc.)
        const numberMatch = this.findSlideNumberMatch(slideIndex + 1); // Convert to 1-based
        if (numberMatch) {
            console.log(`  Found number match: ${numberMatch.filename}`);
            return numberMatch;
        }
        
        // Strategy 3: Keyword match - filename contains words from slide title
        const keywordMatch = this.findKeywordMatch(slideTitle);
        if (keywordMatch) {
            console.log(`  Found keyword match: ${keywordMatch.filename} (score: ${keywordMatch.score})`);
            return keywordMatch;
        }
        
        // No match found - return null (will use placeholder)
        console.log(`  No match found for slide ${slideIndex}`);
        return null;
    }
    
    /**
     * Find exact title match in image filenames
     * @param {string} title - Slide title
     * @returns {object|null} - Image data or null
     */
    findExactTitleMatch(title) {
        if (!title) return null;
        
        // Normalize title for comparison (remove special chars, lowercase)
        const normalizedTitle = title.toLowerCase()
            .replace(/[^a-z0-9\s]/g, '')
            .replace(/\s+/g, '_');
        
        for (const [filename, data] of this.selectedImages) {
            // Get filename without extension
            const nameWithoutExt = filename.substring(0, filename.lastIndexOf('.')) || filename;
            const normalizedFilename = nameWithoutExt.toLowerCase()
                .replace(/[^a-z0-9\s_-]/g, '')
                .replace(/[\s-]+/g, '_');
            
            if (normalizedFilename === normalizedTitle) {
                return {
                    filename: filename,
                    data: data,
                    matchType: 'exact'
                };
            }
        }
        
        return null;
    }
    
    /**
     * Find slide number match in image filenames
     * @param {number} slideNumber - 1-based slide number
     * @returns {object|null} - Image data or null
     */
    findSlideNumberMatch(slideNumber) {
        // Patterns to match: slide1, slide_1, slide-1, slide 1, s1, etc.
        const patterns = [
            `slide${slideNumber}`,
            `slide_${slideNumber}`,
            `slide-${slideNumber}`,
            `slide ${slideNumber}`,
            `s${slideNumber}`,
            `${slideNumber}` // Just the number
        ];
        
        for (const [filename, data] of this.selectedImages) {
            const nameWithoutExt = filename.substring(0, filename.lastIndexOf('.')) || filename;
            const normalizedFilename = nameWithoutExt.toLowerCase();
            
            for (const pattern of patterns) {
                if (normalizedFilename === pattern || 
                    normalizedFilename.startsWith(pattern + '_') ||
                    normalizedFilename.startsWith(pattern + '-') ||
                    normalizedFilename.startsWith(pattern + ' ')) {
                    return {
                        filename: filename,
                        data: data,
                        matchType: 'number'
                    };
                }
            }
        }
        
        return null;
    }
    
    /**
     * Find keyword match in image filenames
     * @param {string} title - Slide title
     * @returns {object|null} - Image data or null
     */
    findKeywordMatch(title) {
        if (!title) return null;
        
        // Extract significant keywords from title (ignore common words)
        const stopWords = ['the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'are', 'was', 'were'];
        const keywords = title.toLowerCase()
            .replace(/[^a-z0-9\s]/g, '')
            .split(/\s+/)
            .filter(word => word.length > 2 && !stopWords.includes(word));
        
        if (keywords.length === 0) return null;
        
        let bestMatch = null;
        let bestScore = 0;
        
        for (const [filename, data] of this.selectedImages) {
            const nameWithoutExt = filename.substring(0, filename.lastIndexOf('.')) || filename;
            const normalizedFilename = nameWithoutExt.toLowerCase();
            
            // Count how many keywords match
            let score = 0;
            for (const keyword of keywords) {
                if (normalizedFilename.includes(keyword)) {
                    score++;
                }
            }
            
            // Update best match if this has a higher score
            if (score > bestScore) {
                bestScore = score;
                bestMatch = {
                    filename: filename,
                    data: data,
                    matchType: 'keyword',
                    score: score
                };
            }
        }
        
        // Only return if at least one keyword matched
        return bestScore > 0 ? bestMatch : null;
    }
    
    /**
     * Get all matched images for slides
     * @param {Array} slides - Array of slide data
     * @returns {Map} - Map of slide index to image data
     */
    getMatchedImagesForSlides(slides) {
        const matchedImages = new Map();
        
        slides.forEach((slide, index) => {
            const match = this.matchImageForSlide(slide, index);
            if (match) {
                matchedImages.set(index, match);
            }
        });
        
        return matchedImages;
    }
    
    /**
     * Check if local images are available
     * @returns {boolean}
     */
    hasImages() {
        return this.selectedImages.size > 0;
    }
    
    /**
     * Get folder information
     * @returns {object} - Folder info
     */
    getFolderInfo() {
        return {
            path: this.folderPath,
            imageCount: this.imageCount,
            hasImages: this.hasImages()
        };
    }
}

// Create and export singleton instance
const localImageHandler = new LocalImageHandler();