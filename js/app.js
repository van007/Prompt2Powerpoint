/**
 * Main application logic for Prompt 2 Powerpoint
 */
class App {
    constructor() {
        // Initialize connection check interval
        this.connectionCheckInterval = null;
        
        // Initialize app
        this.init();
    }

    /**
     * Initialize the application
     */
    async init() {
        // Add event listeners
        this.setupEventListeners();
        
        // Initial connection check
        this.checkConnection();
        
        // Set up connection monitoring
        this.startConnectionMonitoring();
        
        // Initialize PDF.js
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
    }

    /**
     * Set up all event listeners
     */
    setupEventListeners() {
        // Generate presentation event
        document.addEventListener('generate', async (e) => {
            await this.generatePresentation(e.detail);
        });
        
        // Model selected event
        document.addEventListener('modelSelected', (e) => {
            apiClient.setModel(e.detail.modelId);
        });
        
        // Download event
        document.addEventListener('download', async () => {
            await this.downloadPresentation();
        });
        
        // Add slide event
        document.addEventListener('addSlide', async (e) => {
            await this.addSlide(e.detail);
        });
    }

    /**
     * Check connection to the LLM server
     */
    async checkConnection() {
        const providerName = apiClient.provider === 'openrouter' ? 'OpenRouter' : 'Local LLM';
        uiHandler.updateConnectionStatus(`Connecting to ${providerName}...`);
        
        try {
            const connected = await apiClient.checkConnection();
            
            if (connected) {
                uiHandler.updateConnectionStatus(`Connected (${providerName})`);
                
                // Load models
                const models = await apiClient.getModels();
                
                if (models.length > 0) {
                    // Remember current selection before repopulating
                    const currentSelection = apiClient.getSelectedModel();
                    
                    uiHandler.populateModelSelect(models);
                    
                    // Check if previously selected model is still available
                    const modelStillAvailable = models.some(model => model.id === currentSelection);
                    
                    if (currentSelection && modelStillAvailable) {
                        // Restore previous selection
                        document.getElementById('model-select').value = currentSelection;
                    } else if (models[0].id && !currentSelection) {
                        // Only select first model if nothing was previously selected
                        document.getElementById('model-select').value = models[0].id;
                        apiClient.setModel(models[0].id);
                    }
                }
            } else {
                uiHandler.updateConnectionStatus('Disconnected');
            }
        } catch (error) {
            console.error('Connection check error:', error);
            uiHandler.updateConnectionStatus('Disconnected');
        }
    }

    /**
     * Start connection monitoring
     */
    startConnectionMonitoring() {
        // Check connection every 30 seconds
        this.connectionCheckInterval = setInterval(() => {
            this.checkConnection();
        }, 30000);
    }

    /**
     * Stop connection monitoring
     */
    stopConnectionMonitoring() {
        if (this.connectionCheckInterval) {
            clearInterval(this.connectionCheckInterval);
            this.connectionCheckInterval = null;
        }
    }

    /**
     * Generate a presentation based on user input
     * @param {object} inputValues - User input values
     */
    async generatePresentation(inputValues) {
        try {
            // Validate local image selection if needed
            if (inputValues.imageSource === 'local' && !localImageHandler.hasImages()) {
                uiHandler.showError('Please select a folder containing images before generating the presentation.');
                return;
            }
            
            // Disable generate button
            uiHandler.disableGenerateButton();
            
            // Process any uploaded files
            await fileHandler.processAllFiles();
            const contexts = fileHandler.getFileContents();
            
            // Show progress starting at 0%
            uiHandler.updateProgress(0);
            
            // Generate presentation content
            const presentationData = await apiClient.generatePresentation(
                inputValues.prompt,
                contexts,
                inputValues.complexity,
                inputValues.slideCount,
                (progress) => {
                    uiHandler.updateProgress(progress);
                },
                inputValues.imageLayout,
                inputValues.useRealImages,
                inputValues.language || 'en'
            );
            
            // Set the selected theme
            if (inputValues.theme) {
                presentationBuilder.setTheme(inputValues.theme);
            }
            
            // Set selected image layout
            if (inputValues.imageLayout) {
                console.log('Setting selected image layout:', inputValues.imageLayout);
                presentationBuilder.setSelectedImageLayout(inputValues.imageLayout);
            }
            
            // Set image source (new property with backwards compatibility)
            if (inputValues.imageSource) {
                console.log('Setting image source:', inputValues.imageSource);
                presentationBuilder.setImageSource(inputValues.imageSource);
                
                // If using local images, match them to slides
                if (inputValues.imageSource === 'local' && localImageHandler.hasImages()) {
                    console.log('Processing local images for slides...');
                    const matchedImages = localImageHandler.getMatchedImagesForSlides(presentationData.slides);
                    console.log('Matched images:', matchedImages.size, 'images matched to slides');
                    presentationBuilder.setLocalImages(matchedImages);
                }
            } else {
                // Backwards compatibility: use the boolean flag
                presentationBuilder.setUseRealImages(inputValues.useRealImages);
            }
            
            // Set logo if configured
            if (inputValues.logoSettings && inputValues.logoSettings.data) {
                presentationBuilder.setLogo(
                    inputValues.logoSettings.data,
                    inputValues.logoSettings.position,
                    inputValues.logoSettings.size,
                    inputValues.logoSettings.width,
                    inputValues.logoSettings.height
                );
            } else {
                presentationBuilder.clearLogo();
            }
            
            // Initialize presentation builder with the data, original prompt, and language
            console.log('Initializing presentation with data:', presentationData);
            presentationBuilder.initialize(presentationData, inputValues.prompt, inputValues.language || 'en');
            
            // Generate preview
            const previews = presentationBuilder.generatePreviews();
            uiHandler.displayPreview(previews);
            
            // Scroll to results section
            uiHandler.scrollToElement(document.getElementById('result-section'));
            
            // Show success message
            uiHandler.showSuccess('Presentation generated successfully!', 'Success');
        } catch (error) {
            console.error('Generation error:', error);
            uiHandler.showError(error.message || 'Failed to generate presentation');
            
            // Hide progress section
            document.getElementById('progress-section').style.display = 'none';
        } finally {
            // Always re-enable the generate button
            uiHandler.enableGenerateButton();
        }
    }

    /**
     * Download the generated presentation
     */
    async downloadPresentation() {
        try {
            // Show a loading indicator
            uiHandler.showLoading('Preparing your presentation for download...');
            
            // Wait a moment to ensure UI updates
            await new Promise(resolve => setTimeout(resolve, 100));
            
            // Generate and download the presentation
            await presentationBuilder.downloadPresentation();
            
            // Hide loading and show success
            uiHandler.hideLoading();
            uiHandler.showSuccess('Presentation downloaded successfully!', 'Download Complete');
        } catch (error) {
            console.error('Download error:', error);
            uiHandler.hideLoading();
            uiHandler.showError(error.message || 'Failed to download presentation');
        }
    }
    
    /**
     * Add a new slide to the presentation
     * @param {object} slideDetails - Details for the new slide
     */
    async addSlide(slideDetails) {
        try {
            const { title, description, position } = slideDetails;
            
            // Get the current complexity level
            const complexity = presentationBuilder.getCurrentComplexity();
            
            // Get the currently selected image layout
            const imageLayout = uiHandler.getSelectedImageLayout();
            
            // Gather context for better slide generation
            const context = {
                originalPrompt: presentationBuilder.getOriginalPrompt(),
                presentationTitle: presentationBuilder.getPresentationTitle(),
                position: position,
                previousSlide: position > 0 ? presentationBuilder.getSlideAtPosition(position) : null,
                nextSlide: presentationBuilder.getSlideAtPosition(position + 1)
            };
            
            // Get whether to use real images
            const useRealImages = uiHandler.getUseRealImages();
            
            // Get current language setting
            const language = uiHandler.elements.languageSelect ? uiHandler.elements.languageSelect.value : 'en';
            
            // Generate the new slide using the API with context
            const newSlideData = await apiClient.generateSingleSlide(title, description, complexity, context, imageLayout, useRealImages, language);
            
            // Position mapping:
            // Position 0: After title slide = Insert at content index 0 (before first content slide)
            // Position 1: After first content slide = Insert at content index 1 (after first content slide)
            // Position 2: After second content slide = Insert at content index 2 (after second content slide)
            // The position directly corresponds to the content slides array insertion point
            
            console.log(`Inserting slide at content position: ${position}`);
            
            // Insert the slide at the specified position
            presentationBuilder.insertSlide(newSlideData, position);
            
            // Regenerate and display the updated preview
            const updatedPreviews = presentationBuilder.regeneratePreviews();
            uiHandler.displayPreview(updatedPreviews);
            
            // Hide loading and show success
            uiHandler.hideLoading();
            uiHandler.showSuccess(`Slide "${title}" added successfully!`, 'Slide Added');
            
        } catch (error) {
            console.error('Add slide error:', error);
            uiHandler.hideLoading();
            uiHandler.showError(error.message || 'Failed to add slide');
        }
    }
}

// Initialize the app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    const app = new App();
});
