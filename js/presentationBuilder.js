/**
 * Presentation Builder for creating PowerPoint presentations
 */
class PresentationBuilder {
    constructor() {
        this.pptx = null;
        this.presentationData = null;
        this.selectedTheme = 'professional';
        this.originalPrompt = ''; // Store the original prompt
        this.useRealImages = false; // Whether to use real images (local images)
        this.imageCache = new Map(); // Cache for fetched images
        this.localImages = []; // List of available local images
        this.localImageCache = new Map(); // Cache for loaded local images
        
        // Logo properties
        this.logoData = null; // Base64 encoded logo data
        this.logoPosition = 'top-right'; // 'top-right' or 'bottom-left'
        this.logoSize = 'small'; // 'small', 'medium', or 'large'
        this.logoWidth = 0; // Original logo width
        this.logoHeight = 0; // Original logo height
        
        // Define available themes
        this.themes = {
            professional: {
                name: 'Professional',
                headFontFace: 'Arial',
                bodyFontFace: 'Arial',
                primaryColor: '2C3E50',
                secondaryColor: '3498DB',
                textColor: '333333',
                backgroundColor: 'FFFFFF',
                accentColor: 'E74C3C'
            },
            modern: {
                name: 'Modern',
                headFontFace: 'Helvetica',
                bodyFontFace: 'Helvetica',
                primaryColor: '6C63FF',
                secondaryColor: '00D4FF',
                textColor: '2C2C2C',
                backgroundColor: 'F8F9FA',
                accentColor: 'FF6B6B'
            },
            corporate: {
                name: 'Corporate',
                headFontFace: 'Times New Roman',
                bodyFontFace: 'Times New Roman',
                primaryColor: '003366',
                secondaryColor: 'FFD700',
                textColor: '1A1A1A',
                backgroundColor: 'FFFFFF',
                accentColor: '008080'
            },
            creative: {
                name: 'Creative',
                headFontFace: 'Comic Sans MS',
                bodyFontFace: 'Arial',
                primaryColor: 'FF4757',
                secondaryColor: '5F27CD',
                textColor: '2C2C2C',
                backgroundColor: 'FFF3E0',
                accentColor: '00D2D3'
            },
            minimalist: {
                name: 'Minimalist',
                headFontFace: 'Helvetica Neue',
                bodyFontFace: 'Helvetica Neue',
                primaryColor: '000000',
                secondaryColor: '888888',
                textColor: '333333',
                backgroundColor: 'FFFFFF',
                accentColor: 'CCCCCC'
            },
            custom: {
                name: 'Custom',
                headFontFace: 'Helvetica Neue',
                bodyFontFace: 'Helvetica Neue',
                primaryColor: '000000',
                secondaryColor: '888888',
                textColor: '333333',
                backgroundColor: 'FFFFFF',
                accentColor: 'CCCCCC'
            }
        };
        
        // Load saved custom theme if exists
        this.loadCustomTheme();
    }

    /**
     * Scan local images from assets/images folder
     * @returns {Promise<Array>} - Array of image file information
     */
    async scanLocalImages() {
        // Always use the manifest approach since we can't list directories in browser
        console.log('Scanning for local images using manifest.json...');
        const images = await this.scanLocalImagesAlternative();
        console.log('Scan complete. Images found:', images.length);
        return images;
    }

    /**
     * Alternative method to scan local images
     * This tries to load a predefined list or checks for common image names
     * @returns {Promise<Array>} - Array of image file information
     */
    async scanLocalImagesAlternative() {
        // Since we can't dynamically list files in browser, we'll need to either:
        // 1. Maintain a manifest file listing available images
        // 2. Try to load images with common naming patterns
        this.localImages = [];
        
        // Try to load a manifest file if it exists
        try {
            console.log('Attempting to fetch manifest from: assets/images/manifest.json');
            const response = await fetch('assets/images/manifest.json');
            
            if (!response.ok) {
                console.error('Failed to fetch manifest. Status:', response.status);
                return this.localImages;
            }
            
            const manifestText = await response.text();
            console.log('Manifest response received, length:', manifestText.length);
            
            const manifest = JSON.parse(manifestText);
            console.log('Manifest parsed successfully:', manifest);
            
            if (manifest && manifest.images) {
                this.localImages = manifest.images.map(filename => ({
                    filename: filename,
                    path: `assets/images/${filename}`,
                    keywords: this.extractKeywordsFromFilename(filename)
                }));
                console.log(`Successfully loaded ${this.localImages.length} local images from manifest`);
                console.log('Image list:', this.localImages);
                
                // Update UI if images were loaded and toggle is on
                if (this.localImages.length > 0 && typeof uiHandler !== 'undefined' && uiHandler.getUseRealImages()) {
                    uiHandler.updateLocalImagesStatus();
                }
            } else {
                console.error('Manifest does not contain images array:', manifest);
            }
        } catch (error) {
            console.error('Error loading manifest:', error);
            console.error('Full error details:', error.stack);
        }
        
        return this.localImages;
    }

    /**
     * Extract keywords from a filename
     * @param {string} filename - The image filename
     * @returns {Array<string>} - Array of keywords
     */
    extractKeywordsFromFilename(filename) {
        // Remove file extension
        const nameWithoutExt = filename.replace(/\.[^/.]+$/, '');
        
        // Split by common separators and clean up
        const keywords = nameWithoutExt
            .split(/[-_\s]+/)
            .map(word => word.toLowerCase())
            .filter(word => word.length > 2); // Filter out very short words
        
        return keywords;
    }

    /**
     * Extract keywords from text content
     * @param {string} text - Text to extract keywords from
     * @returns {Array<string>} - Array of keywords
     */
    extractKeywords(text) {
        if (!text) return [];
        
        // Common stop words to filter out
        const stopWords = new Set([
            'the', 'and', 'for', 'are', 'with', 'this', 'that', 'have', 'has',
            'will', 'can', 'our', 'your', 'their', 'what', 'when', 'where', 'how',
            'why', 'all', 'would', 'could', 'should', 'may', 'might', 'must',
            'shall', 'will', 'from', 'into', 'through', 'during', 'before', 'after'
        ]);
        
        // Extract words and filter
        const keywords = text
            .toLowerCase()
            .replace(/[^\w\s]/g, ' ') // Remove punctuation
            .split(/\s+/)
            .filter(word => word.length > 2 && !stopWords.has(word));
        
        return [...new Set(keywords)]; // Remove duplicates
    }

    /**
     * Calculate match score between keywords and filename keywords
     * @param {Array<string>} contentKeywords - Keywords from slide content
     * @param {Array<string>} filenameKeywords - Keywords from image filename
     * @returns {number} - Match score (0-100)
     */
    calculateMatchScore(contentKeywords, filenameKeywords) {
        if (!contentKeywords.length || !filenameKeywords.length) return 0;
        
        let matchCount = 0;
        const contentSet = new Set(contentKeywords.map(k => k.toLowerCase()));
        
        for (const keyword of filenameKeywords) {
            if (contentSet.has(keyword.toLowerCase())) {
                matchCount++;
            }
            // Also check for partial matches (e.g., "team" matches "teamwork")
            for (const contentKeyword of contentSet) {
                if (contentKeyword.includes(keyword) || keyword.includes(contentKeyword)) {
                    matchCount += 0.5;
                    break;
                }
            }
        }
        
        // Calculate score as percentage of matched keywords
        const score = (matchCount / Math.max(filenameKeywords.length, contentKeywords.length)) * 100;
        return Math.min(100, score);
    }

    /**
     * Find the best matching local image for a slide
     * @param {object} slideData - The slide data
     * @returns {object|null} - Best matching image or null
     */
    findBestLocalImage(slideData) {
        if (!this.localImages.length) return null;
        
        // Extract keywords from slide content
        const titleKeywords = this.extractKeywords(slideData.title || '');
        const contentKeywords = slideData.content ? 
            this.extractKeywords(Array.isArray(slideData.content) ? 
                slideData.content.join(' ') : slideData.content) : [];
        const descriptionKeywords = this.extractKeywords(slideData.imageDescription || '');
        
        // Combine all keywords with weights
        const allKeywords = [
            ...titleKeywords,
            ...titleKeywords, // Double weight for title
            ...contentKeywords,
            ...descriptionKeywords,
            ...descriptionKeywords // Double weight for image description
        ];
        
        // Score each image
        let bestImage = null;
        let bestScore = 0;
        
        for (const image of this.localImages) {
            const score = this.calculateMatchScore(allKeywords, image.keywords);
            if (score > bestScore) {
                bestScore = score;
                bestImage = image;
            }
        }
        
        // If no good match found (score < 20), return random image
        if (bestScore < 20 && this.localImages.length > 0) {
            const randomIndex = Math.floor(Math.random() * this.localImages.length);
            bestImage = this.localImages[randomIndex];
            console.log(`No good match found for slide "${slideData.title}", using random image`);
        } else if (bestImage) {
            console.log(`Found image "${bestImage.filename}" with score ${bestScore} for slide "${slideData.title}"`);
        }
        
        return bestImage;
    }

    /**
     * Load a local image and convert to base64
     * @param {string} imagePath - Path to the image
     * @returns {Promise<string>} - Base64 encoded image data
     */
    async loadLocalImage(imagePath) {
        // Check cache first
        if (this.localImageCache.has(imagePath)) {
            return this.localImageCache.get(imagePath);
        }
        
        try {
            const response = await fetch(imagePath);
            const blob = await response.blob();
            
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    const base64Data = reader.result;
                    // Cache the result
                    this.localImageCache.set(imagePath, base64Data);
                    resolve(base64Data);
                };
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            console.error(`Failed to load local image ${imagePath}:`, error);
            return null;
        }
    }

    /**
     * Get local image for slide or fallback to placeholder
     * @param {object} slideData - The slide data
     * @returns {Promise<string>} - Image data (base64) or placeholder
     */
    async getLocalImageForSlide(slideData) {
        if (!this.useRealImages || !this.localImages.length) {
            // Return placeholder if not using real images or no images available
            return this.generatePlaceholderPNG();
        }
        
        const bestImage = this.findBestLocalImage(slideData);
        if (bestImage) {
            const imageData = await this.loadLocalImage(bestImage.path);
            if (imageData) {
                return imageData;
            }
        }
        
        // Fallback to placeholder if loading fails
        return this.generatePlaceholderPNG();
    }

    /**
     * Set the theme for the presentation
     * @param {string} themeName - Name of the theme to use
     */
    setTheme(themeName) {
        if (this.themes[themeName]) {
            this.selectedTheme = themeName;
            // Save custom theme when it's selected
            if (themeName === 'custom') {
                this.saveCustomTheme();
            }
        }
    }
    
    /**
     * Get the current theme
     * @returns {object} - Current theme configuration
     */
    getCurrentTheme() {
        return this.themes[this.selectedTheme];
    }
    
    /**
     * Load custom theme from localStorage
     */
    loadCustomTheme() {
        const savedCustomTheme = localStorage.getItem('custom_theme');
        if (savedCustomTheme) {
            try {
                const customTheme = JSON.parse(savedCustomTheme);
                // Merge with default custom theme to ensure all properties exist
                this.themes.custom = {
                    ...this.themes.custom,
                    ...customTheme,
                    name: 'Custom' // Ensure name stays as Custom
                };
            } catch (e) {
                console.error('Error loading custom theme:', e);
            }
        }
    }
    
    /**
     * Save custom theme to localStorage
     */
    saveCustomTheme() {
        if (this.selectedTheme === 'custom') {
            localStorage.setItem('custom_theme', JSON.stringify(this.themes.custom));
        }
    }
    
    /**
     * Update custom theme property
     * @param {string} property - Theme property to update
     * @param {string} value - New value for the property
     */
    updateCustomTheme(property, value) {
        if (this.themes.custom && this.themes.custom.hasOwnProperty(property)) {
            this.themes.custom[property] = value;
            this.saveCustomTheme();
        }
    }
    
    /**
     * Set logo configuration
     * @param {string} logoData - Base64 encoded logo data
     * @param {string} position - Logo position ('top-right' or 'bottom-left')
     * @param {string} size - Logo size ('small', 'medium', or 'large')
     * @param {number} width - Original logo width
     * @param {number} height - Original logo height
     */
    setLogo(logoData, position, size, width = 0, height = 0) {
        this.logoData = logoData;
        this.logoPosition = position || 'top-right';
        this.logoSize = size || 'small';
        this.logoWidth = width;
        this.logoHeight = height;
    }
    
    /**
     * Clear logo configuration
     */
    clearLogo() {
        this.logoData = null;
        this.logoPosition = 'top-right';
        this.logoSize = 'small';
        this.logoWidth = 0;
        this.logoHeight = 0;
    }
    
    /**
     * Add logo to a slide
     * @param {object} slide - PptxGenJS slide object
     */
    addLogoToSlide(slide) {
        if (!this.logoData) {
            return;
        }
        
        // Define height mappings (percentage of slide height)
        const heightMap = {
            'small': 0.05,    // 5% of slide height
            'medium': 0.07,   // 7% of slide height
            'large': 0.10     // 10% of slide height
        };
        
        // Slide aspect ratio (16:9 for widescreen presentations)
        const slideAspectRatio = 16 / 9;
        
        // Fix the height based on size setting
        const heightPercent = heightMap[this.logoSize] * 100;
        
        // Calculate logo aspect ratio
        const logoAspectRatio = this.logoWidth > 0 && this.logoHeight > 0 ? 
            this.logoWidth / this.logoHeight : 1;
        
        // Calculate width to maintain aspect ratio
        // Account for slide aspect ratio when converting height% to width%
        const widthPercent = (heightPercent * logoAspectRatio) / slideAspectRatio;
        
        // Build logo configuration with both width and height
        let logoConfig = {
            data: this.logoData,
            h: `${heightPercent}%`,
            w: `${widthPercent}%`
        };
        
        // Calculate position based on selected option
        if (this.logoPosition === 'top-right') {
            // Position in top-right corner
            logoConfig.x = `${95 - widthPercent}%`;
            logoConfig.y = '5%';
        } else if (this.logoPosition === 'bottom-left') {
            // Position in bottom-left corner
            logoConfig.x = '5%';
            logoConfig.y = `${95 - heightPercent}%`;
        }
        
        // Add the logo image to the slide
        try {
            slide.addImage(logoConfig);
        } catch (error) {
            console.error('Error adding logo to slide:', error);
        }
    }

    /**
     * Initialize a new presentation
     * @param {object} data - Presentation data from the LLM
     * @param {string} prompt - Original prompt used to generate presentation
     */
    initialize(data, prompt = '') {
        this.presentationData = data;
        this.originalPrompt = prompt;
        this.pptx = new PptxGenJS();
        
        // Set presentation properties
        this.pptx.author = 'Prompt 2 Powerpoint';
        this.pptx.company = 'Generated with AI';
        this.pptx.title = data.title || 'AI Generated Presentation';
        
        // Set default layout
        this.pptx.layout = 'LAYOUT_16x9';
        
        // Get current theme
        const theme = this.getCurrentTheme();
        
        // Set theme based on selected option
        this.pptx.theme = {
            headFontFace: theme.headFontFace,
            bodyFontFace: theme.bodyFontFace,
            color: theme.primaryColor,
            background: theme.backgroundColor
        };
    }

    /**
     * Generate the complete presentation
     * @returns {Promise<Blob>} - Presentation as a Blob
     */
    async generatePresentation() {
        if (!this.presentationData || !this.pptx) {
            throw new Error('Presentation not initialized');
        }
        
        try {
            // Create title slide
            this.createTitleSlide(this.presentationData.title);
            
            // Create content slides
            if (this.presentationData.slides && Array.isArray(this.presentationData.slides)) {
                for (const slideData of this.presentationData.slides) {
                    await this.createContentSlide(slideData);
                }
            }
            
            // Create closing slide
            this.createClosingSlide();
            
            // Generate and return the presentation
            return await this.pptx.writeFile({ outputType: 'blob' });
        } catch (error) {
            console.error('Error generating presentation:', error);
            throw error;
        }
    }

    /**
     * Create the title slide
     * @param {string} title - Presentation title
     */
    createTitleSlide(title) {
        const slide = this.pptx.addSlide();
        const theme = this.getCurrentTheme();
        
        // Add logo if configured
        this.addLogoToSlide(slide);
        
        // Add title
        slide.addText(title, {
            x: '10%',
            y: '40%',
            w: '80%',
            fontSize: 44,
            fontFace: theme.headFontFace,
            color: theme.primaryColor,
            bold: true,
            align: 'center'
        });
        
        // Add subtitle
        slide.addText('Generated with Prompt 2 Powerpoint', {
            x: '10%',
            y: '60%',
            w: '80%',
            fontSize: 20,
            fontFace: theme.bodyFontFace,
            color: theme.secondaryColor,
            align: 'center'
        });
        
        // Add date
        const today = new Date();
        const formattedDate = today.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
        
        slide.addText(formattedDate, {
            x: '10%',
            y: '70%',
            w: '80%',
            fontSize: 14,
            fontFace: theme.bodyFontFace,
            color: theme.textColor,
            align: 'center'
        });
    }

    /**
     * Create a content slide
     * @param {object} slideData - Data for this slide
     */
    async createContentSlide(slideData) {
        const slide = this.pptx.addSlide();
        const theme = this.getCurrentTheme();
        
        // Add logo if configured
        this.addLogoToSlide(slide);
        
        // Determine layout type with intelligent fallback
        let layout = slideData.imageLayout || 'none';
        
        // If image layouts are enabled globally but this slide doesn't have one, assign a default
        if (layout === 'none' && this.hasEnabledImageLayouts()) {
            layout = this.getDefaultLayoutForSlide(slideData);
            console.warn(`Slide "${slideData.title}" missing imageLayout, using fallback: ${layout}`);
        }
        
        // Add title (same for all layouts)
        slide.addText(slideData.title || 'Slide Title', {
            x: '5%',
            y: '5%',
            w: '90%',
            h: '15%',
            fontSize: 28,
            fontFace: theme.headFontFace,
            color: theme.primaryColor,
            bold: true
        });
        
        // Handle different layouts
        if (layout === 'full-width') {
            await this.createFullWidthLayout(slide, slideData, theme);
        } else if (layout === 'side-by-side') {
            await this.createSideBySideLayout(slide, slideData, theme);
        } else if (layout === 'text-focus') {
            await this.createTextFocusLayout(slide, slideData, theme);
        } else if (layout === 'background') {
            await this.createBackgroundLayout(slide, slideData, theme);
        } else {
            // Default layout (no image)
            this.createDefaultLayout(slide, slideData, theme);
        }
        
        // Add slide number (same for all layouts)
        const slideNumber = this.pptx.slides.length;
        slide.addText(`${slideNumber}`, {
            x: '90%',
            y: '95%',
            w: '5%',
            fontSize: 12,
            fontFace: theme.bodyFontFace,
            color: theme.secondaryColor,
            align: 'right'
        });
        
        // Add notes if available
        if (slideData.notes) {
            slide.addNotes(slideData.notes);
        }
    }
    
    /**
     * Create default layout (no image)
     */
    createDefaultLayout(slide, slideData, theme) {
        // Add content/bullet points
        if (slideData.content) {
            if (Array.isArray(slideData.content) && slideData.content.length > 0) {
                // Format as bullet points
                const bulletPoints = slideData.content.map(point => ({ text: String(point) }));
                
                slide.addText(bulletPoints, {
                    x: '5%',
                    y: '25%',
                    w: '90%',
                    h: '65%',
                    fontSize: 18,
                    fontFace: theme.bodyFontFace,
                    color: theme.textColor,
                    bullet: { type: 'bullet' },
                    lineSpacing: 28
                });
            } else if (typeof slideData.content === 'string') {
                // Add as paragraph text
                slide.addText(slideData.content, {
                    x: '5%',
                    y: '25%',
                    w: '90%',
                    h: '65%',
                    fontSize: 18,
                    fontFace: theme.bodyFontFace,
                    color: theme.textColor,
                    lineSpacing: 28
                });
            } else {
                // Try to convert to string
                try {
                    const contentStr = typeof slideData.content === 'object' ? 
                        JSON.stringify(slideData.content) : String(slideData.content);
                        
                    slide.addText(contentStr, {
                        x: '5%',
                        y: '25%',
                        w: '90%',
                        h: '65%',
                        fontSize: 18,
                        fontFace: theme.bodyFontFace,
                        color: theme.textColor,
                        lineSpacing: 28
                    });
                } catch (e) {
                    console.error('Error adding slide content:', e);
                    // Add fallback text
                    slide.addText('Content unavailable', {
                        x: '5%',
                        y: '25%',
                        w: '90%',
                        h: '65%',
                        fontSize: 18,
                        fontFace: theme.bodyFontFace,
                        color: theme.secondaryColor,
                        lineSpacing: 28
                    });
                }
            }
        }
    }
    
    /**
     * Generate transparent PNG placeholder image as base64
     * @returns {string} - Base64 encoded transparent PNG
     */
    generatePlaceholderPNG() {
        // This is a 1x1 transparent PNG
        // In PowerPoint, this will appear as an empty image placeholder that users can easily replace
        const transparentPNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
        return `data:image/png;base64,${transparentPNG}`;
    }
    
    /**
     * Add image placeholder to slide
     * @param {object} slide - PptxGenJS slide object
     * @param {object} options - Placement options
     */
    addImagePlaceholder(slide, options) {
        const defaults = {
            x: '5%',
            y: '25%',
            w: '90%',
            h: '50%',
            altText: 'Image placeholder - click to replace'
        };
        
        const config = { ...defaults, ...options };
        
        // Generate transparent placeholder PNG
        const placeholderData = this.generatePlaceholderPNG();
        
        // Add image to slide with placeholder properties
        slide.addImage({
            data: placeholderData,
            x: config.x,
            y: config.y,
            w: config.w,
            h: config.h,
            altText: config.altText,
            placeholder: true  // This might help PowerPoint recognize it as a placeholder
        });
        
        // Add text in the center to indicate where the image should go
        // This provides visual feedback without blocking the image placeholder
        slide.addText(config.altText, {
            x: config.x,
            y: config.y,
            w: config.w,
            h: config.h,
            fontSize: 14,
            color: '666666',
            align: 'center',
            valign: 'middle',
            italic: true
        });
    }
    
    /**
     * Create full-width layout (image at top, text below)
     */
    async createFullWidthLayout(slide, slideData, theme) {
        const imageOptions = {
            x: '5%',
            y: '22%',
            w: '90%',
            h: '35%'
        };
        
        // Get local image or placeholder
        const imageData = await this.getLocalImageForSlide(slideData);
        
        if (this.useRealImages && imageData !== this.generatePlaceholderPNG()) {
            // Add local image
            slide.addImage({
                data: imageData,
                x: imageOptions.x,
                y: imageOptions.y,
                w: imageOptions.w,
                h: imageOptions.h,
                altText: slideData.imageDescription || 'Image from local assets'
            });
        } else {
            // Use placeholder
            this.addImagePlaceholder(slide, {
                ...imageOptions,
                altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
            });
        }
        
        // Add content below image
        if (slideData.content && Array.isArray(slideData.content)) {
            const bulletPoints = slideData.content.map(point => ({ text: String(point) }));
            slide.addText(bulletPoints, {
                x: '5%',
                y: '60%',
                w: '90%',
                h: '30%',
                fontSize: 16,
                fontFace: theme.bodyFontFace,
                color: theme.textColor,
                bullet: { type: 'bullet' },
                lineSpacing: 24
            });
        }
    }
    
    /**
     * Create side-by-side layout (image left, text right)
     */
    async createSideBySideLayout(slide, slideData, theme) {
        const imageOptions = {
            x: '5%',
            y: '22%',
            w: '42%',
            h: '65%'
        };
        
        // Get local image or placeholder
        const imageData = await this.getLocalImageForSlide(slideData);
        
        if (this.useRealImages && imageData !== this.generatePlaceholderPNG()) {
            // Add local image
            slide.addImage({
                data: imageData,
                x: imageOptions.x,
                y: imageOptions.y,
                w: imageOptions.w,
                h: imageOptions.h,
                altText: slideData.imageDescription || 'Image from local assets'
            });
        } else {
            // Use placeholder
            this.addImagePlaceholder(slide, {
                ...imageOptions,
                altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
            });
        }
        
        // Add content on right
        if (slideData.content && Array.isArray(slideData.content)) {
            const bulletPoints = slideData.content.map(point => ({ text: String(point) }));
            slide.addText(bulletPoints, {
                x: '52%',
                y: '25%',
                w: '43%',
                h: '60%',
                fontSize: 18,
                fontFace: theme.bodyFontFace,
                color: theme.textColor,
                bullet: { type: 'bullet' },
                lineSpacing: 28
            });
        }
    }
    
    /**
     * Create text-focus layout (small image right, more text space)
     */
    async createTextFocusLayout(slide, slideData, theme) {
        // Add content on left (larger space)
        if (slideData.content && Array.isArray(slideData.content)) {
            const bulletPoints = slideData.content.map(point => ({ text: String(point) }));
            slide.addText(bulletPoints, {
                x: '5%',
                y: '25%',
                w: '60%',
                h: '65%',
                fontSize: 18,
                fontFace: theme.bodyFontFace,
                color: theme.textColor,
                bullet: { type: 'bullet' },
                lineSpacing: 28
            });
        }
        
        // Add small image on right
        const imageOptions = {
            x: '70%',
            y: '25%',
            w: '25%',
            h: '35%'
        };
        
        // Get local image or placeholder
        const imageData = await this.getLocalImageForSlide(slideData);
        
        if (this.useRealImages && imageData !== this.generatePlaceholderPNG()) {
            // Add local image
            slide.addImage({
                data: imageData,
                x: imageOptions.x,
                y: imageOptions.y,
                w: imageOptions.w,
                h: imageOptions.h,
                altText: slideData.imageDescription || 'Image from local assets'
            });
        } else {
            // Use placeholder
            this.addImagePlaceholder(slide, {
                ...imageOptions,
                altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
            });
        }
    }
    
    /**
     * Create background layout (full background image with text overlay)
     */
    async createBackgroundLayout(slide, slideData, theme) {
        const imageOptions = {
            x: '0%',
            y: '0%',
            w: '100%',
            h: '100%'
        };
        
        // Get local image or placeholder
        const imageData = await this.getLocalImageForSlide(slideData);
        
        if (this.useRealImages && imageData !== this.generatePlaceholderPNG()) {
            // Add local image as background
            slide.addImage({
                data: imageData,
                x: imageOptions.x,
                y: imageOptions.y,
                w: imageOptions.w,
                h: imageOptions.h,
                altText: slideData.imageDescription || 'Background image from local assets'
            });
        } else {
            // Use placeholder
            this.addImagePlaceholder(slide, {
                ...imageOptions,
                altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
            });
        }
        
        // Add semi-transparent background for title
        slide.addShape(this.pptx.ShapeType.rect, {
            x: '8%',
            y: '32%',
            w: '84%',
            h: '20%',
            fill: { color: theme.backgroundColor, alpha: 50 },
            line: { type: 'none' },
            shadow: {
                type: 'outer',
                color: '000000',
                blur: 8,
                offset: 2,
                angle: 180,
                opacity: 0.2
            }
        });
        
        // Re-add title with better contrast
        slide.addText(slideData.title || 'Slide Title', {
            x: '10%',
            y: '35%',
            w: '80%',
            h: '15%',
            fontSize: 32,
            fontFace: theme.headFontFace,
            color: theme.primaryColor,
            bold: true,
            align: 'center',
            shadow: {
                type: 'outer',
                color: '000000',
                blur: 3,
                offset: 2,
                angle: 45,
                opacity: 0.5
            }
        });
        
        // Add content centered
        if (slideData.content && Array.isArray(slideData.content)) {
            // Add semi-transparent background for content
            slide.addShape(this.pptx.ShapeType.rect, {
                x: '13%',
                y: '53%',
                w: '74%',
                h: '25%',
                fill: { color: theme.backgroundColor, alpha: 50 },
                line: { type: 'none' },
                shadow: {
                    type: 'outer',
                    color: '000000',
                    blur: 8,
                    offset: 2,
                    angle: 180,
                    opacity: 0.2
                }
            });
            
            const bulletPoints = slideData.content.map(point => ({ text: String(point) }));
            slide.addText(bulletPoints, {
                x: '15%',
                y: '55%',
                w: '70%',
                h: '20%',
                fontSize: 20,
                fontFace: theme.bodyFontFace,
                color: theme.textColor,
                bullet: { type: 'bullet' },
                lineSpacing: 30,
                align: 'center',
                shadow: {
                    type: 'outer',
                    color: '000000',
                    blur: 2,
                    offset: 1,
                    angle: 45,
                    opacity: 0.4
                }
            });
        }
    }

    /**
     * Create the closing slide
     */
    createClosingSlide() {
        const slide = this.pptx.addSlide();
        const theme = this.getCurrentTheme();
        
        // Add logo if configured
        this.addLogoToSlide(slide);
        
        // Add thank you message
        slide.addText('Thank You', {
            x: '10%',
            y: '40%',
            w: '80%',
            fontSize: 44,
            fontFace: theme.headFontFace,
            color: theme.primaryColor,
            bold: true,
            align: 'center'
        });
        
        // Add generated by message
        slide.addText('Generated by Prompt 2 Powerpoint', {
            x: '10%',
            y: '60%',
            w: '80%',
            fontSize: 16,
            fontFace: theme.bodyFontFace,
            color: theme.secondaryColor,
            align: 'center'
        });
    }

    /**
     * Generate a preview of the slides
     * @returns {Array<object>} - Array of slide preview data
     */
    generatePreviews() {
        if (!this.presentationData) {
            return [];
        }
        
        const previews = [];
        
        // Title slide
        previews.push({
            title: this.presentationData.title || 'Title Slide',
            content: ['Generated with Prompt 2 Powerpoint'],
            notes: ''
        });
        
        // Content slides
        if (this.presentationData.slides && Array.isArray(this.presentationData.slides)) {
            for (const slideData of this.presentationData.slides) {
                previews.push({
                    title: slideData.title || 'Slide',
                    content: slideData.content || ['Content'],
                    notes: slideData.notes || ''
                });
            }
        }
        
        // Closing slide
        previews.push({
            title: 'Thank You',
            content: ['Generated by Prompt 2 Powerpoint'],
            notes: ''
        });
        
        return previews;
    }

    /**
     * Download the presentation
     */
    async downloadPresentation() {
        if (!this.presentationData || !this.pptx) {
            throw new Error('Presentation not initialized');
        }
        
        try {
            console.log('Starting presentation download...');
            console.log('Presentation data:', this.presentationData);
            
            // Completely regenerate the presentation to ensure all slides are included
            // First reset the existing presentation
            this.pptx = new PptxGenJS();
            
            // Set presentation properties again
            this.pptx.author = 'Prompt 2 Powerpoint';
            this.pptx.company = 'Generated with AI';
            this.pptx.title = this.presentationData.title || 'AI Generated Presentation';
            this.pptx.layout = 'LAYOUT_16x9';
            
            // Get current theme
            const theme = this.getCurrentTheme();
            
            // Set theme based on selected option
            this.pptx.theme = {
                headFontFace: theme.headFontFace,
                bodyFontFace: theme.bodyFontFace,
                color: theme.primaryColor,
                background: theme.backgroundColor
            };
            
            // Recreate title slide
            this.createTitleSlide(this.presentationData.title);
            
            // Recreate content slides
            if (this.presentationData.slides && Array.isArray(this.presentationData.slides)) {
                console.log(`Creating ${this.presentationData.slides.length} content slides`);
                for (const slideData of this.presentationData.slides) {
                    await this.createContentSlide(slideData);
                }
            } else {
                console.warn('No slides array found or slides is not an array:', this.presentationData.slides);
            }
            
            // Create closing slide
            this.createClosingSlide();
            
            // Sanitize the title for the filename
            const sanitizedTitle = this.presentationData.title
                .replace(/[^a-z0-9]/gi, '_')
                .toLowerCase();
            
            const filename = `${sanitizedTitle}_prompt_2_powerpoint.pptx`;
            console.log(`Saving presentation as: ${filename}`);
            
            // Write the file and trigger download
            await this.pptx.writeFile({ fileName: filename });
            
            console.log('Presentation download complete');
            return true;
        } catch (error) {
            console.error('Error during presentation download:', error);
            throw error;
        }
    }
    
    /**
     * Insert a new slide at a specific position
     * @param {object} slideData - The slide data to insert
     * @param {number} position - Position to insert the slide (0-based index)
     */
    insertSlide(slideData, position) {
        if (!this.presentationData) {
            throw new Error('Presentation not initialized');
        }
        
        // Ensure slides array exists
        if (!this.presentationData.slides || !Array.isArray(this.presentationData.slides)) {
            this.presentationData.slides = [];
        }
        
        // Insert the slide at the specified position
        this.presentationData.slides.splice(position, 0, slideData);
        
        console.log(`Inserted slide at position ${position}:`, slideData);
    }
    
    /**
     * Get the current complexity level from the most recent generation
     * @returns {string} - The complexity level
     */
    getCurrentComplexity() {
        // Try to get from UI if available
        const complexitySelect = document.getElementById('complexity-select');
        if (complexitySelect && complexitySelect.value) {
            return complexitySelect.value;
        }
        
        // Default to standard
        return 'standard';
    }
    
    /**
     * Get the current slides data for preview generation
     * @returns {Array<object>} - Array of slide data
     */
    getCurrentSlidesData() {
        if (!this.presentationData || !this.presentationData.slides) {
            return [];
        }
        
        return this.presentationData.slides;
    }
    
    /**
     * Regenerate previews after slide insertion
     * @returns {Array<object>} - Updated array of slide preview data
     */
    regeneratePreviews() {
        return this.generatePreviews();
    }
    
    /**
     * Get the presentation title
     * @returns {string} - Presentation title
     */
    getPresentationTitle() {
        return this.presentationData ? this.presentationData.title : '';
    }
    
    /**
     * Get the original prompt
     * @returns {string} - Original prompt
     */
    getOriginalPrompt() {
        return this.originalPrompt;
    }
    
    /**
     * Check if any image layouts are enabled
     * @returns {boolean}
     */
    hasEnabledImageLayouts() {
        // Check if we have stored image layout preferences
        // This would be set during initialization from the UI
        return this.imageLayoutsEnabled || false;
    }
    
    /**
     * Set the selected image layout
     * @param {string} layout - The selected layout type
     */
    setSelectedImageLayout(layout) {
        this.selectedImageLayout = layout;
        this.imageLayoutsEnabled = layout && layout !== 'none';
    }
    
    /**
     * Get default layout for slides
     * @param {object} slideData - Slide data (not used with single layout)
     * @returns {string} - Layout type
     */
    getDefaultLayoutForSlide(slideData) {
        // With single layout selection, always return the selected layout
        return this.selectedImageLayout || 'none';
    }
    
    /**
     * Get slide at specific position (including title and closing slides)
     * @param {number} position - Position in the full presentation (0 = title slide)
     * @returns {object|null} - Slide data or null if position is invalid
     */
    getSlideAtPosition(position) {
        const previews = this.generatePreviews();
        if (position >= 0 && position < previews.length) {
            return previews[position];
        }
        return null;
    }
    
    /**
     * Set whether to use real images from local assets
     * @param {boolean} useReal - Whether to use real images
     */
    setUseRealImages(useReal) {
        this.useRealImages = useReal;
    }
}

// Create and export a singleton instance
const presentationBuilder = new PresentationBuilder();
