/**
 * UI Handler for managing user interface interactions
 */
class UiHandler {
    constructor() {
        // DOM elements
        this.elements = {
            promptInput: document.getElementById('prompt-input'),
            generateBtn: document.getElementById('generate-btn'),
            complexitySelect: document.getElementById('complexity-select'),
            slidesSelect: document.getElementById('slides-select'),
            themeSelect: document.getElementById('theme-select'),
            uploadArea: document.getElementById('upload-area'),
            fileInput: document.getElementById('file-input'),
            browseBtn: document.getElementById('browse-btn'),
            fileList: document.getElementById('file-list'),
            modelSelect: document.getElementById('model-select'),
            connectionStatus: document.getElementById('connection-status'),
            progressSection: document.getElementById('progress-section'),
            progressBar: document.getElementById('progress-bar'),
            progressPercentage: document.getElementById('progress-percentage'),
            resultSection: document.getElementById('result-section'),
            presentationPreview: document.getElementById('presentation-preview'),
            downloadBtn: document.getElementById('download-btn'),
            
            // Theme preview elements
            themePreview: document.getElementById('theme-preview'),
            headingFontPreview: document.getElementById('heading-font-preview'),
            bodyFontPreview: document.getElementById('body-font-preview'),
            primaryColorSwatch: document.getElementById('primary-color-swatch'),
            secondaryColorSwatch: document.getElementById('secondary-color-swatch'),
            textColorSwatch: document.getElementById('text-color-swatch'),
            backgroundColorSwatch: document.getElementById('background-color-swatch'),
            accentColorSwatch: document.getElementById('accent-color-swatch'),
            
            // Settings elements
            settingsBtn: document.getElementById('settings-btn'),
            urlSettingsModal: document.getElementById('url-settings-modal'),
            closeModalBtn: document.getElementById('close-modal'),
            urlForm: document.getElementById('url-form'),
            baseUrlInput: document.getElementById('base-url-input'),
            resetUrlBtn: document.getElementById('reset-url-btn'),
            
            // Provider settings elements
            providerLocal: document.getElementById('provider-local'),
            providerOpenRouter: document.getElementById('provider-openrouter'),
            localSettings: document.getElementById('local-settings'),
            openRouterSettings: document.getElementById('openrouter-settings'),
            apiKeyInput: document.getElementById('api-key-input'),
            
            // Image layout elements
            layoutNone: document.getElementById('layout-none'),
            layoutFullWidth: document.getElementById('layout-full-width'),
            layoutSideBySide: document.getElementById('layout-side-by-side'),
            layoutTextFocus: document.getElementById('layout-text-focus'),
            layoutBackground: document.getElementById('layout-background'),
            
            // Image source elements
            sourcePlaceholder: document.getElementById('source-placeholder'),
            sourcePexels: document.getElementById('source-pexels'),
            sourceLocal: document.getElementById('source-local'),
            localFolderSelection: document.getElementById('local-folder-selection'),
            selectFolderBtn: document.getElementById('select-folder-btn'),
            folderInput: document.getElementById('folder-input'),
            folderInfo: document.getElementById('folder-info'),
            folderName: document.getElementById('folder-name'),
            folderImageCount: document.getElementById('folder-image-count'),
            clearFolderBtn: document.getElementById('clear-folder-btn'),
            
            // Legacy: Real images toggle (for backwards compatibility)
            useRealImagesToggle: document.getElementById('use-real-images'),
            imageTypeDescription: document.getElementById('image-type-description'),
            placeholderDesc: document.querySelector('.placeholder-desc'),
            realImagesDesc: document.querySelector('.real-images-desc'),
            
            // Pexels API
            pexelsApiKeyInput: document.getElementById('pexels-api-key-input'),
            
            // Font selects
            headingFontSelect: document.getElementById('heading-font-select'),
            bodyFontSelect: document.getElementById('body-font-select'),
            
            // Color inputs
            primaryColorInput: document.getElementById('primary-color-input'),
            secondaryColorInput: document.getElementById('secondary-color-input'),
            textColorInput: document.getElementById('text-color-input'),
            backgroundColorInput: document.getElementById('background-color-input'),
            accentColorInput: document.getElementById('accent-color-input'),
            
            // Logo elements
            logoUploadArea: document.getElementById('logo-upload-area'),
            logoUploadContent: document.getElementById('logo-upload-content'),
            logoFileInput: document.getElementById('logo-file-input'),
            logoPreviewContainer: document.getElementById('logo-preview-container'),
            logoPreviewImage: document.getElementById('logo-preview-image'),
            removeLogoBtn: document.getElementById('remove-logo-btn')
        };
        
        // Logo data storage
        this.logoData = null;
        this.logoPosition = 'top-right';
        this.logoSize = 'small';
        this.logoWidth = 0;
        this.logoHeight = 0;
        
        // Model select state tracking
        this.isModelSelectFocused = false;
        this.isModelSelectOpen = false;
        
        // Event handlers
        this.setupEventListeners();
        
        // Initialize URL field with current value
        this.initializeUrlField();
        
        // Initialize logo settings
        this.initializeLogoSettings();
        
        // Store the insertion position for new slides
        this.slideInsertPosition = 0;
    }

    /**
     * Set up all event listeners
     */
    setupEventListeners() {
        // Generate button
        this.elements.generateBtn.addEventListener('click', () => {
            this.onGenerateClicked();
        });
        
        // File upload via button
        this.elements.browseBtn.addEventListener('click', () => {
            this.elements.fileInput.click();
        });
        
        // File input change
        this.elements.fileInput.addEventListener('change', (e) => {
            this.handleFileUpload(e.target.files);
        });
        
        // Drag and drop
        this.elements.uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.elements.uploadArea.classList.add('dragover');
        });
        
        this.elements.uploadArea.addEventListener('dragleave', () => {
            this.elements.uploadArea.classList.remove('dragover');
        });
        
        this.elements.uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            this.elements.uploadArea.classList.remove('dragover');
            this.handleFileUpload(e.dataTransfer.files);
        });
        
        // Download button
        this.elements.downloadBtn.addEventListener('click', () => {
            this.onDownloadClicked();
        });
        
        // Model selection
        this.elements.modelSelect.addEventListener('change', () => {
            this.onModelSelected();
        });
        
        // Track model select focus state
        this.elements.modelSelect.addEventListener('focus', () => {
            this.isModelSelectFocused = true;
        });
        
        this.elements.modelSelect.addEventListener('blur', () => {
            this.isModelSelectFocused = false;
            this.isModelSelectOpen = false;
            
            // Apply pending model update if exists
            if (this.pendingModelUpdate) {
                setTimeout(() => {
                    // Use timeout to ensure blur has completed
                    this.populateModelSelect(this.pendingModelUpdate, true);
                }, 100);
            }
        });
        
        // Track when dropdown is actually opened (mousedown on select)
        this.elements.modelSelect.addEventListener('mousedown', () => {
            // If already focused, this click is opening/closing the dropdown
            if (this.isModelSelectFocused) {
                this.isModelSelectOpen = !this.isModelSelectOpen;
            } else {
                // First click focuses and opens
                this.isModelSelectOpen = true;
            }
        });
        
        // Theme selection
        this.elements.themeSelect.addEventListener('change', () => {
            this.onThemeSelected();
        });
        
        // Enter key in prompt input
        this.elements.promptInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.onGenerateClicked();
            }
        });
        
        // Settings button - open modal
        this.elements.settingsBtn.addEventListener('click', () => {
            this.openUrlSettingsModal();
        });
        
        // Close modal button
        this.elements.closeModalBtn.addEventListener('click', () => {
            this.closeUrlSettingsModal();
        });
        
        // Click outside modal to close
        this.elements.urlSettingsModal.addEventListener('click', (e) => {
            if (e.target === this.elements.urlSettingsModal) {
                this.closeUrlSettingsModal();
            }
        });
        
        // URL form submission
        this.elements.urlForm.addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveBaseUrl();
        });
        
        // Reset URL button
        this.elements.resetUrlBtn.addEventListener('click', () => {
            this.resetBaseUrl();
        });
        
        // Provider radio buttons
        this.elements.providerLocal.addEventListener('change', () => {
            this.onProviderChanged();
        });
        
        this.elements.providerOpenRouter.addEventListener('change', () => {
            this.onProviderChanged();
        });
        
        // Add slide modal handlers
        this.setupAddSlideModalHandlers();
        
        // Image source selection handlers
        this.setupImageSourceHandlers();
        
        // Legacy: Real images toggle (for backwards compatibility)
        if (this.elements.useRealImagesToggle) {
            this.elements.useRealImagesToggle.addEventListener('change', () => {
                this.onImageTypeToggled();
            });
        }
        
        // Initialize image source state
        this.initializeImageSource();
        
        // Custom theme event listeners
        this.setupCustomThemeListeners();
    }
    
    /**
     * Initialize the URL input field with the current base URL
     */
    initializeUrlField() {
        // Get current URL from API client
        const currentUrl = apiClient.baseUrl;
        if (currentUrl) {
            this.elements.baseUrlInput.value = currentUrl;
        }
        
        // Set provider radio button
        if (apiClient.provider === 'openrouter') {
            this.elements.providerOpenRouter.checked = true;
        } else {
            this.elements.providerLocal.checked = true;
        }
        
        // Set API key if available
        if (apiClient.openRouterApiKey) {
            this.elements.apiKeyInput.value = apiClient.openRouterApiKey;
        }
        
        // Set Pexels API key if available
        const savedPexelsKey = pexelsClient.loadSavedApiKey();
        if (savedPexelsKey && this.elements.pexelsApiKeyInput) {
            this.elements.pexelsApiKeyInput.value = savedPexelsKey;
        }
        
        // Show/hide appropriate settings
        this.onProviderChanged();
    }
    
    /**
     * Handle provider change
     */
    onProviderChanged() {
        const isOpenRouter = this.elements.providerOpenRouter.checked;
        
        if (isOpenRouter) {
            this.elements.localSettings.style.display = 'none';
            this.elements.openRouterSettings.style.display = 'block';
        } else {
            this.elements.localSettings.style.display = 'block';
            this.elements.openRouterSettings.style.display = 'none';
        }
    }
    
    /**
     * Open the URL settings modal
     */
    openUrlSettingsModal() {
        this.elements.urlSettingsModal.classList.add('active');
    }
    
    /**
     * Close the URL settings modal
     */
    closeUrlSettingsModal() {
        this.elements.urlSettingsModal.classList.remove('active');
    }
    
    /**
     * Save the base URL and update the API client
     */
    async saveBaseUrl() {
        // Get provider selection
        const isOpenRouter = this.elements.providerOpenRouter.checked;
        const provider = isOpenRouter ? 'openrouter' : 'local';
        
        // Update provider
        apiClient.setProvider(provider);
        
        // Validate based on provider
        if (provider === 'local') {
            const newUrl = this.elements.baseUrlInput.value.trim();
            
            if (!newUrl) {
                this.showError('Please enter a valid URL');
                return;
            }
            
            // Update the base URL for local provider
            await apiClient.setBaseUrl(newUrl);
        } else {
            // OpenRouter
            const apiKey = this.elements.apiKeyInput.value.trim();
            
            if (!apiKey) {
                this.showError('Please enter your OpenRouter API key');
                return;
            }
            
            // Update API key
            apiClient.setOpenRouterApiKey(apiKey);
        }
        
        // Save Pexels API key if provided
        const pexelsApiKey = this.elements.pexelsApiKeyInput.value.trim();
        if (pexelsApiKey) {
            pexelsClient.setApiKey(pexelsApiKey);
        }
        
        // Show loading
        this.showLoading(`Connecting to ${provider === 'openrouter' ? 'OpenRouter' : 'LLM server'}...`);
        
        try {
            // Test connection
            const success = await apiClient.checkConnection();
            
            this.hideLoading();
            
            if (success) {
                this.showSuccess(`Connected to ${provider === 'openrouter' ? 'OpenRouter' : 'LLM server'} successfully`);
                this.closeUrlSettingsModal();
                
                // Update connection status
                this.updateConnectionStatus('Connected');
                
                // Reload models
                const models = await apiClient.getModels();
                if (models.length > 0) {
                    // Remember current selection before repopulating
                    const currentSelection = apiClient.getSelectedModel();
                    
                    this.populateModelSelect(models);
                    
                    // Restore selection if still available
                    if (currentSelection && models.some(m => m.id === currentSelection)) {
                        this.elements.modelSelect.value = currentSelection;
                    }
                }
            } else {
                this.showError(`Failed to connect to ${provider === 'openrouter' ? 'OpenRouter' : 'LLM server'}. Please check your settings and try again.`);
            }
        } catch (error) {
            this.hideLoading();
            this.showError(`Error connecting to ${provider === 'openrouter' ? 'OpenRouter' : 'LLM server'}: ` + error.message);
        }
    }
    
    /**
     * Reset the base URL to default
     */
    async resetBaseUrl() {
        // Reset to local provider with default URL
        this.elements.providerLocal.checked = true;
        this.elements.baseUrlInput.value = 'http://127.0.0.1:1234';
        this.elements.apiKeyInput.value = '';
        
        // Show local settings
        this.onProviderChanged();
        
        // Save the default settings
        await this.saveBaseUrl();
    }

    /**
     * Handle file upload
     * @param {FileList} files - Uploaded files
     */
    handleFileUpload(files) {
        if (!files || files.length === 0) return;
        
        const newFiles = fileHandler.addFiles(files);
        this.updateFileList();
        
        // Reset file input
        this.elements.fileInput.value = '';
    }

    /**
     * Update the file list display
     */
    updateFileList() {
        const files = fileHandler.getFiles();
        this.elements.fileList.innerHTML = '';
        
        files.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            
            const fileIcon = file.type === 'application/pdf' ? 'fa-file-pdf' : 'fa-file-lines';
            
            fileItem.innerHTML = `
                <div class="file-name">
                    <i class="fas ${fileIcon}"></i>
                    <span>${file.name}</span>
                </div>
                <button class="remove-file" data-index="${index}">
                    <i class="fas fa-times"></i>
                </button>
            `;
            
            // Add remove button handler
            const removeBtn = fileItem.querySelector('.remove-file');
            removeBtn.addEventListener('click', () => {
                fileHandler.removeFile(index);
                this.updateFileList();
            });
            
            this.elements.fileList.appendChild(fileItem);
        });
    }

    /**
     * Populate model selection dropdown
     * @param {Array} models - Available models
     * @param {boolean} forceUpdate - Force update even if dropdown is focused/open
     */
    populateModelSelect(models, forceUpdate = false) {
        const select = this.elements.modelSelect;
        
        // Skip update if dropdown is currently focused/open and not forcing update
        if (!forceUpdate && (this.isModelSelectFocused || this.isModelSelectOpen)) {
            // Store the models for later update when dropdown is closed
            this.pendingModelUpdate = models;
            return;
        }
        
        // Check if models have actually changed
        const currentOptions = Array.from(select.options).slice(1); // Skip the first "Select Model" option
        const currentModelIds = currentOptions.map(opt => opt.value);
        const newModelIds = models.map(model => model.id);
        
        // If models haven't changed, don't update
        if (currentModelIds.length === newModelIds.length && 
            currentModelIds.every((id, index) => id === newModelIds[index])) {
            return;
        }
        
        // Remember current selection
        const currentSelection = select.value;
        
        // Clear existing options
        select.innerHTML = '<option value="" disabled selected>Select Model</option>';
        
        // Add new options
        models.forEach(model => {
            const option = document.createElement('option');
            option.value = model.id;
            option.textContent = model.id;
            select.appendChild(option);
        });
        
        // Restore selection if it still exists
        if (currentSelection && models.some(m => m.id === currentSelection)) {
            select.value = currentSelection;
        }
        
        // Enable the select
        select.disabled = false;
        
        // Clear pending update
        this.pendingModelUpdate = null;
    }

    /**
     * Update connection status display
     * @param {string} status - Connection status
     */
    updateConnectionStatus(status) {
        const element = this.elements.connectionStatus;
        element.textContent = status;
        element.className = '';
        
        switch (status.toLowerCase()) {
            case 'connected':
                element.classList.add('connected');
                break;
            case 'disconnected':
                element.classList.add('disconnected');
                break;
            case 'connecting':
                element.classList.add('connecting');
                break;
        }
    }

    /**
     * Update progress display
     * @param {number} percentage - Progress percentage (0-100)
     */
    updateProgress(percentage) {
        this.elements.progressBar.style.width = `${percentage}%`;
        this.elements.progressPercentage.textContent = `${percentage}%`;
        
        // Show progress section if not already visible
        this.elements.progressSection.style.display = 'block';
        
        // If this is the first update (0%), scroll to the progress section
        if (percentage === 0) {
            this.scrollToElement(this.elements.progressSection);
        }
        
        // If complete, hide after a delay
        if (percentage >= 100) {
            setTimeout(() => {
                this.elements.progressSection.style.display = 'none';
            }, 1000);
        }
    }
    
    /**
     * Disable the generate button during processing
     */
    disableGenerateButton() {
        this.elements.generateBtn.disabled = true;
        this.elements.generateBtn.classList.add('disabled');
    }
    
    /**
     * Enable the generate button when processing is complete
     */
    enableGenerateButton() {
        this.elements.generateBtn.disabled = false;
        this.elements.generateBtn.classList.remove('disabled');
    }
    
    /**
     * Scroll to a specific element
     * @param {HTMLElement} element - Element to scroll to
     */
    scrollToElement(element) {
        if (element) {
            element.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }

    /**
     * Display presentation preview
     * @param {Array} slides - Slide preview data
     */
    displayPreview(slides) {
        const container = this.elements.presentationPreview;
        container.innerHTML = '';
        
        slides.forEach((slide, index) => {
            // Create slide row container
            const slideRow = document.createElement('div');
            slideRow.className = 'slide-row';
            
            // Add plus icon after title slide (index 0) to insert first content slide
            if (index === 0) {
                // This is the title slide - add plus icon after it
                const afterTitleIcon = this.createPlusIcon(0); // Position 0 in content slides
                slideRow.appendChild(afterTitleIcon);
            }
            
            // Create slide preview
            const slidePreview = document.createElement('div');
            slidePreview.className = 'slide-preview';
            
            // Create a more detailed preview with actual content
            const slideContent = document.createElement('div');
            slideContent.className = 'slide-content';
            
            // Add slide title
            const slideTitle = document.createElement('div');
            slideTitle.className = 'slide-title';
            slideTitle.textContent = slide.title || `Slide ${index + 1}`;
            
            // Add slide content (first bullet point or truncated)
            const contentPreview = document.createElement('div');
            contentPreview.className = 'slide-content-preview';
            
            if (slide.content && slide.content.length > 0) {
                // Show the first content item, truncated if needed
                const firstItem = slide.content[0];
                contentPreview.textContent = firstItem.length > 30 ? 
                    firstItem.substring(0, 30) + '...' : 
                    firstItem;
                    
                // Add a count if there are more items
                if (slide.content.length > 1) {
                    const itemCount = document.createElement('div');
                    itemCount.className = 'item-count';
                    itemCount.textContent = `+${slide.content.length - 1} more`;
                    contentPreview.appendChild(itemCount);
                }
            }
            
            // Add slide number
            const slideNumber = document.createElement('div');
            slideNumber.className = 'slide-number';
            slideNumber.textContent = `${index + 1}`;
            
            // Assemble the preview
            slideContent.appendChild(slideTitle);
            slideContent.appendChild(contentPreview);
            
            slidePreview.appendChild(slideContent);
            slidePreview.appendChild(slideNumber);
            
            // Add a tooltip with more complete slide info
            slidePreview.title = `${slide.title}\n${Array.isArray(slide.content) ? slide.content.join('\n') : slide.content || ''}`;
            
            // Add click handler to show full slide details
            slidePreview.addEventListener('click', () => {
                this.showSlideDetails(slide, index + 1);
            });
            
            // Add slide to row
            slideRow.appendChild(slidePreview);
            
            // Add plus icon after content slides (not after title or closing slide)
            if (index > 0 && index < slides.length - 1) {
                // This is a content slide - add plus icon after it
                // Position in content slides array = index - 1 (content slide position) + 1 (after position)
                const contentPosition = index - 1 + 1; // = index
                const afterContentIcon = this.createPlusIcon(contentPosition);
                slideRow.appendChild(afterContentIcon);
            }
            
            container.appendChild(slideRow);
        });
        
        // Show result section
        this.elements.resultSection.style.display = 'block';
    }

    /**
     * Show an error message
     * @param {string} message - Error message
     * @param {string} title - Error title
     */
    showError(message, title = 'Error') {
        Swal.fire({
            title: title,
            text: message,
            icon: 'error',
            confirmButtonText: 'OK',
            confirmButtonColor: '#2D5BFF'
        });
    }

    /**
     * Show a success message
     * @param {string} message - Success message
     * @param {string} title - Success title
     */
    showSuccess(message, title = 'Success') {
        Swal.fire({
            title: title,
            text: message,
            icon: 'success',
            confirmButtonText: 'OK',
            confirmButtonColor: '#2D5BFF'
        });
    }

    /**
     * Show a loading indicator
     * @param {string} message - Loading message to display
     */
    showLoading(message = 'Loading...') {
        Swal.fire({
            title: message,
            allowOutsideClick: false,
            didOpen: () => {
                Swal.showLoading();
            }
        });
    }

    /**
     * Hide the loading indicator
     */
    hideLoading() {
        Swal.close();
    }

    /**
     * Validate input fields before generation
     * @returns {boolean} - Whether inputs are valid
     */
    validateInputs() {
        const prompt = this.elements.promptInput.value.trim();
        const complexity = this.elements.complexitySelect.value;
        const slideCount = this.elements.slidesSelect.value;
        const modelId = this.elements.modelSelect.value;
        const theme = this.elements.themeSelect.value;
        
        if (!prompt) {
            this.showError('Please enter a presentation prompt');
            return false;
        }
        
        if (!complexity) {
            this.showError('Please select a complexity level');
            return false;
        }
        
        if (!theme) {
            this.showError('Please select a presentation theme');
            return false;
        }
        
        if (!slideCount) {
            this.showError('Please select the number of slides');
            return false;
        }
        
        if (!modelId) {
            this.showError('Please select a model');
            return false;
        }
        
        return true;
    }

    /**
     * Get user input values
     * @returns {object} - User input values
     */
    getInputValues() {
        return {
            prompt: this.elements.promptInput.value.trim(),
            complexity: this.elements.complexitySelect.value,
            slideCount: parseInt(this.elements.slidesSelect.value, 10),
            modelId: this.elements.modelSelect.value,
            theme: this.elements.themeSelect.value,
            imageLayout: this.getSelectedImageLayout(),
            useRealImages: this.getUseRealImages(), // Backwards compatibility
            imageSource: this.getImageSource(), // New property
            logoSettings: this.getLogoSettings()
        };
    }
    
    /**
     * Get the selected image layout
     * @returns {string} - Selected layout type or 'none'
     */
    getSelectedImageLayout() {
        // Find the checked radio button
        const checkedRadio = document.querySelector('input[name="image-layout"]:checked');
        return checkedRadio ? checkedRadio.value : 'none';
    }

    /**
     * Event handler for generate button
     */
    onGenerateClicked() {
        if (this.validateInputs()) {
            const event = new CustomEvent('generate', {
                detail: this.getInputValues()
            });
            
            document.dispatchEvent(event);
        }
    }

    /**
     * Event handler for model selection
     */
    onModelSelected() {
        const modelId = this.elements.modelSelect.value;
        
        if (modelId) {
            const event = new CustomEvent('modelSelected', {
                detail: { modelId }
            });
            
            document.dispatchEvent(event);
        }
    }
    
    /**
     * Setup image source selection handlers
     */
    setupImageSourceHandlers() {
        // Image source radio buttons
        const sourceRadios = document.querySelectorAll('input[name="image-source"]');
        sourceRadios.forEach(radio => {
            radio.addEventListener('change', () => {
                this.onImageSourceChanged();
            });
        });
        
        // Folder selection button
        if (this.elements.selectFolderBtn) {
            this.elements.selectFolderBtn.addEventListener('click', () => {
                this.elements.folderInput.click();
            });
        }
        
        // Folder input change
        if (this.elements.folderInput) {
            this.elements.folderInput.addEventListener('change', async (e) => {
                await this.handleFolderSelection(e.target.files);
            });
        }
        
        // Clear folder button
        if (this.elements.clearFolderBtn) {
            this.elements.clearFolderBtn.addEventListener('click', () => {
                this.clearFolderSelection();
            });
        }
    }
    
    /**
     * Initialize image source state
     */
    initializeImageSource() {
        // Check saved image source preference
        const savedSource = localStorage.getItem('image_source');
        if (savedSource) {
            const sourceRadio = document.querySelector(`input[name="image-source"][value="${savedSource}"]`);
            if (sourceRadio) {
                sourceRadio.checked = true;
            }
        }
        
        // Trigger change handler to update UI
        this.onImageSourceChanged();
    }
    
    /**
     * Handle image source change
     */
    onImageSourceChanged() {
        const selectedSource = this.getImageSource();
        
        // Save preference
        localStorage.setItem('image_source', selectedSource);
        
        // Show/hide local folder selection
        if (this.elements.localFolderSelection) {
            this.elements.localFolderSelection.style.display = 
                selectedSource === 'local' ? 'block' : 'none';
        }
        
        // Check API key requirements
        if (selectedSource === 'pexels' && !pexelsClient.hasApiKey()) {
            // Show warning but don't force change
            this.showWarning('Pexels API key not configured. Please add it in settings to use stock images.');
        }
    }
    
    /**
     * Handle folder selection
     * @param {FileList} files - Selected files from folder
     */
    async handleFolderSelection(files) {
        if (!files || files.length === 0) return;
        
        this.showLoading('Processing folder...');
        
        try {
            const result = await localImageHandler.processFolder(files);
            
            if (result.success) {
                // Update UI with folder info
                if (this.elements.folderName) {
                    this.elements.folderName.textContent = result.folderPath || 'Selected folder';
                }
                if (this.elements.folderImageCount) {
                    this.elements.folderImageCount.textContent = `${result.imageCount} images`;
                }
                if (this.elements.folderInfo) {
                    this.elements.folderInfo.style.display = 'flex';
                }
                
                this.hideLoading();
                this.showSuccess(`Found ${result.imageCount} images in folder`);
            } else {
                this.hideLoading();
                this.showError(result.message || 'Failed to process folder');
            }
        } catch (error) {
            this.hideLoading();
            this.showError('Error processing folder: ' + error.message);
        }
        
        // Reset file input
        this.elements.folderInput.value = '';
    }
    
    /**
     * Clear folder selection
     */
    clearFolderSelection() {
        localImageHandler.clear();
        
        if (this.elements.folderInfo) {
            this.elements.folderInfo.style.display = 'none';
        }
        if (this.elements.folderName) {
            this.elements.folderName.textContent = 'No folder selected';
        }
        if (this.elements.folderImageCount) {
            this.elements.folderImageCount.textContent = '0 images';
        }
        
        this.showInfo('Folder selection cleared');
    }
    
    /**
     * Get the selected image source
     * @returns {string} - 'placeholder', 'pexels', or 'local'
     */
    getImageSource() {
        const checkedRadio = document.querySelector('input[name="image-source"]:checked');
        return checkedRadio ? checkedRadio.value : 'placeholder';
    }
    
    /**
     * Event handler for theme selection
     */
    onThemeSelected() {
        const themeName = this.elements.themeSelect.value;
        
        if (themeName) {
            this.updateThemePreview(themeName);
        } else {
            this.elements.themePreview.style.display = 'none';
        }
    }
    
    /**
     * Update theme preview with selected theme
     * @param {string} themeName - Name of the selected theme
     */
    updateThemePreview(themeName) {
        const themes = presentationBuilder.themes;
        const theme = themes[themeName];
        
        if (!theme) return;
        
        // Show the preview
        this.elements.themePreview.style.display = 'block';
        
        const isCustomTheme = themeName === 'custom';
        
        // Handle font display
        if (isCustomTheme) {
            // Hide font preview spans and show selects
            this.elements.headingFontPreview.style.display = 'none';
            this.elements.bodyFontPreview.style.display = 'none';
            this.elements.headingFontSelect.style.display = 'block';
            this.elements.bodyFontSelect.style.display = 'block';
            
            // Set select values
            this.elements.headingFontSelect.value = theme.headFontFace;
            this.elements.bodyFontSelect.value = theme.bodyFontFace;
            
            // Update select font styles
            this.elements.headingFontSelect.style.fontFamily = theme.headFontFace;
            this.elements.bodyFontSelect.style.fontFamily = theme.bodyFontFace;
        } else {
            // Show font preview spans and hide selects
            this.elements.headingFontPreview.style.display = 'inline-block';
            this.elements.bodyFontPreview.style.display = 'inline-block';
            this.elements.headingFontSelect.style.display = 'none';
            this.elements.bodyFontSelect.style.display = 'none';
            
            // Update font previews
            this.elements.headingFontPreview.style.fontFamily = theme.headFontFace;
            this.elements.headingFontPreview.textContent = `${theme.headFontFace}`;
            this.elements.bodyFontPreview.style.fontFamily = theme.bodyFontFace;
            this.elements.bodyFontPreview.textContent = `${theme.bodyFontFace}`;
        }
        
        // Update color swatches
        this.elements.primaryColorSwatch.style.backgroundColor = `#${theme.primaryColor}`;
        this.elements.secondaryColorSwatch.style.backgroundColor = `#${theme.secondaryColor}`;
        this.elements.textColorSwatch.style.backgroundColor = `#${theme.textColor}`;
        this.elements.backgroundColorSwatch.style.backgroundColor = `#${theme.backgroundColor}`;
        this.elements.accentColorSwatch.style.backgroundColor = `#${theme.accentColor}`;
        
        // Make color swatches clickable for custom theme
        const colorSwatches = document.querySelectorAll('.color-swatch');
        if (isCustomTheme) {
            colorSwatches.forEach(swatch => {
                swatch.style.cursor = 'pointer';
                swatch.classList.add('clickable');
            });
            
            // Set color input values
            this.elements.primaryColorInput.value = `#${theme.primaryColor}`;
            this.elements.secondaryColorInput.value = `#${theme.secondaryColor}`;
            this.elements.textColorInput.value = `#${theme.textColor}`;
            this.elements.backgroundColorInput.value = `#${theme.backgroundColor}`;
            this.elements.accentColorInput.value = `#${theme.accentColor}`;
        } else {
            colorSwatches.forEach(swatch => {
                swatch.style.cursor = 'default';
                swatch.classList.remove('clickable');
            });
        }
    }

    /**
     * Event handler for download button
     */
    onDownloadClicked() {
        const event = new CustomEvent('download');
        document.dispatchEvent(event);
    }

    /**
     * Set up event listeners for add slide modal
     */
    setupAddSlideModalHandlers() {
        const addSlideModal = document.getElementById('add-slide-modal');
        const closeAddSlideModal = document.getElementById('close-add-slide-modal');
        const cancelAddSlideBtn = document.getElementById('cancel-add-slide-btn');
        const addSlideForm = document.getElementById('add-slide-form');
        
        // Close modal button
        closeAddSlideModal.addEventListener('click', () => {
            this.closeAddSlideModal();
        });
        
        // Cancel button
        cancelAddSlideBtn.addEventListener('click', () => {
            this.closeAddSlideModal();
        });
        
        // Click outside modal to close
        addSlideModal.addEventListener('click', (e) => {
            if (e.target === addSlideModal) {
                this.closeAddSlideModal();
            }
        });
        
        // Form submission
        addSlideForm.addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleAddSlideSubmit();
        });
    }
    
    /**
     * Create a plus icon for adding slides
     * @param {number} position - Position where the slide should be inserted
     * @returns {HTMLElement} - Plus icon element
     */
    createPlusIcon(position) {
        const plusIcon = document.createElement('div');
        plusIcon.className = 'plus-icon';
        plusIcon.innerHTML = '<i class="fas fa-plus"></i>';
        plusIcon.title = 'Add new slide';
        
        plusIcon.addEventListener('click', () => {
            this.openAddSlideModal(position);
        });
        
        return plusIcon;
    }
    
    /**
     * Open the add slide modal
     * @param {number} position - Position to insert the slide
     */
    openAddSlideModal(position) {
        this.slideInsertPosition = position;
        const modal = document.getElementById('add-slide-modal');
        
        // Clear form fields
        document.getElementById('slide-title-input').value = '';
        document.getElementById('slide-description-input').value = '';
        
        // Show modal
        modal.classList.add('active');
        
        // Focus on title input
        document.getElementById('slide-title-input').focus();
    }
    
    /**
     * Close the add slide modal
     */
    closeAddSlideModal() {
        const modal = document.getElementById('add-slide-modal');
        modal.classList.remove('active');
    }
    
    /**
     * Handle add slide form submission
     */
    async handleAddSlideSubmit() {
        const titleInput = document.getElementById('slide-title-input');
        const descriptionInput = document.getElementById('slide-description-input');
        
        const title = titleInput.value.trim();
        const description = descriptionInput.value.trim();
        
        if (!title) {
            this.showError('Please enter a slide title');
            return;
        }
        
        try {
            // Close modal first
            this.closeAddSlideModal();
            
            // Show loading
            this.showLoading('Generating new slide...');
            
            // Trigger slide creation event
            const event = new CustomEvent('addSlide', {
                detail: {
                    title,
                    description,
                    position: this.slideInsertPosition
                }
            });
            
            document.dispatchEvent(event);
            
        } catch (error) {
            this.hideLoading();
            this.showError('Error creating slide: ' + error.message);
        }
    }
    
    /**
     * Legacy: Handle image type toggle change (backwards compatibility)
     */
    onImageTypeToggled() {
        // This is a legacy method for backwards compatibility
        // The new implementation uses image source selection
        const useRealImages = this.elements.useRealImagesToggle ? this.elements.useRealImagesToggle.checked : false;
        
        // Map to new image source system
        if (useRealImages) {
            // Check if Pexels API key is configured
            if (!pexelsClient.hasApiKey()) {
                this.showError('Please configure your Pexels API key in settings to use real images', 'API Key Required');
                // Reset toggle
                if (this.elements.useRealImagesToggle) {
                    this.elements.useRealImagesToggle.checked = false;
                }
            } else {
                // Set to pexels source
                if (this.elements.sourcePexels) {
                    this.elements.sourcePexels.checked = true;
                    this.onImageSourceChanged();
                }
            }
        } else {
            // Set to placeholder source
            if (this.elements.sourcePlaceholder) {
                this.elements.sourcePlaceholder.checked = true;
                this.onImageSourceChanged();
            }
        }
    }
    
    /**
     * Legacy: Get whether to use real images (backwards compatibility)
     * @returns {boolean}
     */
    getUseRealImages() {
        // Map from new image source to boolean for backwards compatibility
        const source = this.getImageSource();
        return source === 'pexels';
    }
    
    /**
     * Show detailed slide information in a modal
     * @param {object} slide - Slide data
     * @param {number} slideNumber - Slide number
     */
    showSlideDetails(slide, slideNumber) {
        // Format content items as bullet points
        let contentHtml = '<p>No content</p>';
        
        if (slide.content) {
            if (Array.isArray(slide.content) && slide.content.length > 0) {
                contentHtml = `<ul>${slide.content.map(item => `<li>${item}</li>`).join('')}</ul>`;
            } else if (typeof slide.content === 'string') {
                contentHtml = `<p>${slide.content}</p>`;
            } else if (typeof slide.content === 'object') {
                contentHtml = `<p>${JSON.stringify(slide.content)}</p>`;
            }
        }
        
        // Show slide details in a SweetAlert modal
        Swal.fire({
            title: `Slide ${slideNumber}: ${slide.title}`,
            html: `
                <div class="slide-details">
                    <h4>Content:</h4>
                    ${contentHtml}
                    ${slide.notes ? `<h4>Notes:</h4><p>${slide.notes}</p>` : ''}
                </div>
            `,
            confirmButtonText: 'Close',
            confirmButtonColor: '#2D5BFF',
            width: '600px'
        });
    }
    
    /**
     * Set up event listeners for custom theme controls
     */
    setupCustomThemeListeners() {
        // Font controls
        if (this.elements.headingFontSelect) {
            this.elements.headingFontSelect.addEventListener('change', (e) => {
                this.updateCustomThemeProperty('headFontFace', e.target.value);
                e.target.style.fontFamily = e.target.value;
            });
        }
        
        if (this.elements.bodyFontSelect) {
            this.elements.bodyFontSelect.addEventListener('change', (e) => {
                this.updateCustomThemeProperty('bodyFontFace', e.target.value);
                e.target.style.fontFamily = e.target.value;
            });
        }
        
        // Color swatch click handlers
        if (this.elements.primaryColorSwatch) {
            this.elements.primaryColorSwatch.addEventListener('click', () => {
                if (this.elements.themeSelect.value === 'custom') {
                    this.elements.primaryColorInput.click();
                }
            });
        }
        
        if (this.elements.secondaryColorSwatch) {
            this.elements.secondaryColorSwatch.addEventListener('click', () => {
                if (this.elements.themeSelect.value === 'custom') {
                    this.elements.secondaryColorInput.click();
                }
            });
        }
        
        if (this.elements.textColorSwatch) {
            this.elements.textColorSwatch.addEventListener('click', () => {
                if (this.elements.themeSelect.value === 'custom') {
                    this.elements.textColorInput.click();
                }
            });
        }
        
        if (this.elements.backgroundColorSwatch) {
            this.elements.backgroundColorSwatch.addEventListener('click', () => {
                if (this.elements.themeSelect.value === 'custom') {
                    this.elements.backgroundColorInput.click();
                }
            });
        }
        
        if (this.elements.accentColorSwatch) {
            this.elements.accentColorSwatch.addEventListener('click', () => {
                if (this.elements.themeSelect.value === 'custom') {
                    this.elements.accentColorInput.click();
                }
            });
        }
        
        // Color input change handlers
        if (this.elements.primaryColorInput) {
            this.elements.primaryColorInput.addEventListener('input', (e) => {
                this.updateCustomThemeProperty('primaryColor', e.target.value.substring(1));
                this.elements.primaryColorSwatch.style.backgroundColor = e.target.value;
            });
        }
        
        if (this.elements.secondaryColorInput) {
            this.elements.secondaryColorInput.addEventListener('input', (e) => {
                this.updateCustomThemeProperty('secondaryColor', e.target.value.substring(1));
                this.elements.secondaryColorSwatch.style.backgroundColor = e.target.value;
            });
        }
        
        if (this.elements.textColorInput) {
            this.elements.textColorInput.addEventListener('input', (e) => {
                this.updateCustomThemeProperty('textColor', e.target.value.substring(1));
                this.elements.textColorSwatch.style.backgroundColor = e.target.value;
            });
        }
        
        if (this.elements.backgroundColorInput) {
            this.elements.backgroundColorInput.addEventListener('input', (e) => {
                this.updateCustomThemeProperty('backgroundColor', e.target.value.substring(1));
                this.elements.backgroundColorSwatch.style.backgroundColor = e.target.value;
            });
        }
        
        if (this.elements.accentColorInput) {
            this.elements.accentColorInput.addEventListener('input', (e) => {
                this.updateCustomThemeProperty('accentColor', e.target.value.substring(1));
                this.elements.accentColorSwatch.style.backgroundColor = e.target.value;
            });
        }
    }
    
    /**
     * Update custom theme property
     * @param {string} property - Property to update
     * @param {string} value - New value
     */
    updateCustomThemeProperty(property, value) {
        // Update the theme
        presentationBuilder.updateCustomTheme(property, value);
    }
    
    /**
     * Initialize logo settings
     */
    initializeLogoSettings() {
        // Load saved logo settings
        const savedLogoSettings = localStorage.getItem('logoSettings');
        if (savedLogoSettings) {
            try {
                const settings = JSON.parse(savedLogoSettings);
                this.logoData = settings.data || null;
                this.logoPosition = settings.position || 'top-right';
                this.logoSize = settings.size || 'small';
                this.logoWidth = settings.width || 0;
                this.logoHeight = settings.height || 0;
                
                // Update UI to reflect saved settings
                if (this.logoData) {
                    this.showLogoPreview(this.logoData);
                }
                
                // Set radio buttons
                document.querySelectorAll('input[name="logo-position"]').forEach(radio => {
                    radio.checked = radio.value === this.logoPosition;
                });
                document.querySelectorAll('input[name="logo-size"]').forEach(radio => {
                    radio.checked = radio.value === this.logoSize;
                });
            } catch (e) {
                console.error('Error loading logo settings:', e);
            }
        }
        
        // Set up logo upload event listeners
        this.setupLogoEventListeners();
    }
    
    /**
     * Set up logo-related event listeners
     */
    setupLogoEventListeners() {
        // Logo upload click
        if (this.elements.logoUploadContent) {
            this.elements.logoUploadContent.addEventListener('click', () => {
                this.elements.logoFileInput.click();
            });
        }
        
        // Logo file input change
        if (this.elements.logoFileInput) {
            this.elements.logoFileInput.addEventListener('change', (e) => {
                this.handleLogoUpload(e);
            });
        }
        
        // Remove logo button
        if (this.elements.removeLogoBtn) {
            this.elements.removeLogoBtn.addEventListener('click', () => {
                this.removeLogo();
            });
        }
        
        // Logo position radio buttons
        document.querySelectorAll('input[name="logo-position"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.logoPosition = e.target.value;
                this.saveLogoSettings();
            });
        });
        
        // Logo size radio buttons
        document.querySelectorAll('input[name="logo-size"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.logoSize = e.target.value;
                this.saveLogoSettings();
            });
        });
    }
    
    /**
     * Handle logo file upload
     * @param {Event} event - File input change event
     */
    handleLogoUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        // Validate file type
        const validTypes = ['image/png', 'image/jpeg', 'image/jpg'];
        if (!validTypes.includes(file.type)) {
            this.showError('Please upload a PNG or JPG image');
            return;
        }
        
        // Validate file size (max 5MB)
        const maxSize = 5 * 1024 * 1024;
        if (file.size > maxSize) {
            this.showError('Logo file size must be less than 5MB');
            return;
        }
        
        // Read file as base64
        const reader = new FileReader();
        reader.onload = (e) => {
            this.logoData = e.target.result;
            
            // Create an image to get dimensions
            const img = new Image();
            img.onload = () => {
                this.logoWidth = img.width;
                this.logoHeight = img.height;
                this.showLogoPreview(this.logoData);
                this.saveLogoSettings();
            };
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    }
    
    /**
     * Show logo preview
     * @param {string} logoData - Base64 encoded logo data
     */
    showLogoPreview(logoData) {
        if (this.elements.logoPreviewImage) {
            this.elements.logoPreviewImage.src = logoData;
        }
        if (this.elements.logoUploadContent) {
            this.elements.logoUploadContent.style.display = 'none';
        }
        if (this.elements.logoPreviewContainer) {
            this.elements.logoPreviewContainer.style.display = 'flex';
        }
    }
    
    /**
     * Remove logo
     */
    removeLogo() {
        this.logoData = null;
        this.logoWidth = 0;
        this.logoHeight = 0;
        
        // Reset UI
        if (this.elements.logoUploadContent) {
            this.elements.logoUploadContent.style.display = 'flex';
        }
        if (this.elements.logoPreviewContainer) {
            this.elements.logoPreviewContainer.style.display = 'none';
        }
        if (this.elements.logoFileInput) {
            this.elements.logoFileInput.value = '';
        }
        
        // Clear saved settings
        this.saveLogoSettings();
    }
    
    /**
     * Save logo settings to localStorage
     */
    saveLogoSettings() {
        const settings = {
            data: this.logoData,
            position: this.logoPosition,
            size: this.logoSize,
            width: this.logoWidth,
            height: this.logoHeight
        };
        localStorage.setItem('logoSettings', JSON.stringify(settings));
    }
    
    /**
     * Get current logo settings
     * @returns {object} Logo settings
     */
    getLogoSettings() {
        return {
            data: this.logoData,
            position: this.logoPosition,
            size: this.logoSize,
            width: this.logoWidth,
            height: this.logoHeight
        };
    }
    
    /**
     * Show warning message
     * @param {string} message - Warning message
     */
    showWarning(message) {
        Swal.fire({
            icon: 'warning',
            title: 'Warning',
            text: message,
            confirmButtonText: 'OK'
        });
    }
    
    /**
     * Show info message
     * @param {string} message - Info message
     */
    showInfo(message) {
        Swal.fire({
            icon: 'info',
            title: 'Info',
            text: message,
            timer: 2000,
            showConfirmButton: false
        });
    }
}

// Create and export a singleton instance
const uiHandler = new UiHandler();
