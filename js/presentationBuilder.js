/**
 * Presentation Builder for creating PowerPoint presentations
 */
class PresentationBuilder {
    constructor() {
        this.pptx = null;
        this.presentationData = null;
        this.selectedTheme = 'professional';
        this.originalPrompt = ''; // Store the original prompt
        this.useRealImages = false; // Whether to use real images from Pexels
        this.imageCache = new Map(); // Cache for fetched images
        
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
            }
        };
    }

    /**
     * Set the theme for the presentation
     * @param {string} themeName - Name of the theme to use
     */
    setTheme(themeName) {
        if (this.themes[themeName]) {
            this.selectedTheme = themeName;
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
        
        // Try to use real image if enabled
        if (this.useRealImages) {
            const imageData = await this.fetchImageForSlide(slideData, 'full-width');
            if (imageData) {
                await this.addRealImage(slide, imageData, imageOptions);
            } else {
                // Fallback to placeholder
                this.addImagePlaceholder(slide, {
                    ...imageOptions,
                    altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
                });
            }
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
        
        // Try to use real image if enabled
        if (this.useRealImages) {
            const imageData = await this.fetchImageForSlide(slideData, 'side-by-side');
            if (imageData) {
                await this.addRealImage(slide, imageData, imageOptions);
            } else {
                this.addImagePlaceholder(slide, {
                    ...imageOptions,
                    altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
                });
            }
        } else {
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
        
        if (this.useRealImages) {
            const imageData = await this.fetchImageForSlide(slideData, 'text-focus');
            if (imageData) {
                await this.addRealImage(slide, imageData, imageOptions);
            } else {
                this.addImagePlaceholder(slide, {
                    ...imageOptions,
                    altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
                });
            }
        } else {
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
        
        // Try to use real image if enabled
        if (this.useRealImages) {
            const imageData = await this.fetchImageForSlide(slideData, 'background');
            if (imageData) {
                await this.addRealImage(slide, imageData, imageOptions);
            } else {
                this.addImagePlaceholder(slide, {
                    ...imageOptions,
                    altText: slideData.imageDescription || 'Right Click -> Change Picture -> Choose Option to Replace'
                });
            }
        } else {
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
     * Set whether to use real images from Pexels
     * @param {boolean} useReal - Whether to use real images
     */
    setUseRealImages(useReal) {
        this.useRealImages = useReal;
    }
    
    /**
     * Add real image from Pexels to slide
     * @param {object} slide - PptxGenJS slide object
     * @param {object} imageData - Image data from Pexels
     * @param {object} options - Placement options
     */
    async addRealImage(slide, imageData, options) {
        const defaults = {
            x: '5%',
            y: '25%',
            w: '90%',
            h: '50%'
        };
        
        const config = { ...defaults, ...options };
        
        try {
            // Add the real image
            slide.addImage({
                path: imageData.url,
                x: config.x,
                y: config.y,
                w: config.w,
                h: config.h,
                altText: imageData.alt || 'Image from Pexels'
            });
            
            // Add attribution to speaker notes
            const currentNotes = slide._slideObjects.filter(obj => obj._type === 'notes')[0];
            const existingNotes = currentNotes ? currentNotes.text : '';
            const attribution = `\n\n${imageData.attribution.text}`;
            
            if (!existingNotes.includes(attribution)) {
                slide.addNotes(existingNotes + attribution);
            }
            
            return true;
        } catch (error) {
            console.error('Error adding real image:', error);
            // Fallback to placeholder
            this.addImagePlaceholder(slide, options);
            return false;
        }
    }
    
    /**
     * Fetch and cache image for slide
     * @param {object} slideData - Slide data containing image description
     * @param {string} layout - Layout type
     * @returns {Promise<object|null>} - Image data or null
     */
    async fetchImageForSlide(slideData, layout) {
        if (!pexelsClient.hasApiKey()) {
            return null;
        }
        
        const cacheKey = `${slideData.title}-${layout}`;
        
        // Check cache first
        if (this.imageCache.has(cacheKey)) {
            return this.imageCache.get(cacheKey);
        }
        
        try {
            // Use image description if available, otherwise use slide title and content
            const searchContext = slideData.imageDescription || 
                                `${slideData.title} ${slideData.content ? slideData.content[0] : ''}`;
            
            // Get theme colors for better image matching
            const themeColors = this.getCurrentTheme();
            
            const imageData = await pexelsClient.getImageForSlide(searchContext, layout, themeColors);
            
            if (imageData) {
                // Cache the result
                this.imageCache.set(cacheKey, imageData);
                return imageData;
            }
        } catch (error) {
            console.error('Error fetching image for slide:', error);
        }
        
        return null;
    }
}

// Create and export a singleton instance
const presentationBuilder = new PresentationBuilder();
