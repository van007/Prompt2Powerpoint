/**
 * API client for local LLM communication
 */
class ApiClient {
    constructor() {
        // Default settings
        this.provider = 'local'; // 'local' or 'openrouter'
        this.baseUrl = 'http://127.0.0.1:1234';
        this.openRouterApiKey = '';
        this.openRouterBaseUrl = 'https://openrouter.ai/api/v1';
        
        this.connectionStatus = 'disconnected';
        this.availableModels = [];
        this.selectedModel = null;
        this.abortController = null;
        
        // Load saved settings from localStorage
        this.loadSavedSettings();
    }
    
    /**
     * Set a new base URL for the API
     * @param {string} url - New base URL
     * @returns {Promise<boolean>} - Whether connection was successful with new URL
     */
    async setBaseUrl(url) {
        // Validate URL format
        try {
            // Check if URL is valid
            new URL(url);
            
            // Update the base URL
            this.baseUrl = url;
            
            // Save to localStorage
            localStorage.setItem('llm_base_url', url);
            
            // Test connection with new URL
            return await this.checkConnection();
        } catch (error) {
            console.error('Invalid URL format:', error);
            return false;
        }
    }
    
    /**
     * Set the API provider
     * @param {string} provider - 'local' or 'openrouter'
     */
    setProvider(provider) {
        this.provider = provider;
        localStorage.setItem('llm_provider', provider);
        
        // Load saved model for this provider
        const modelKey = `selected_model_${provider}`;
        const savedModel = localStorage.getItem(modelKey);
        if (savedModel) {
            this.selectedModel = savedModel;
        } else {
            this.selectedModel = null;
        }
    }
    
    /**
     * Set OpenRouter API key
     * @param {string} apiKey - OpenRouter API key
     */
    setOpenRouterApiKey(apiKey) {
        this.openRouterApiKey = apiKey;
        localStorage.setItem('openrouter_api_key', apiKey);
    }
    
    /**
     * Load saved settings from localStorage
     */
    loadSavedSettings() {
        // Load provider
        const savedProvider = localStorage.getItem('llm_provider');
        if (savedProvider) {
            this.provider = savedProvider;
        }
        
        // Load local URL
        const savedUrl = localStorage.getItem('llm_base_url');
        if (savedUrl) {
            this.baseUrl = savedUrl;
        }
        
        // Load OpenRouter API key
        const savedApiKey = localStorage.getItem('openrouter_api_key');
        if (savedApiKey) {
            this.openRouterApiKey = savedApiKey;
        }
        
        // Load selected model for current provider
        const modelKey = `selected_model_${this.provider}`;
        const savedModel = localStorage.getItem(modelKey);
        if (savedModel) {
            this.selectedModel = savedModel;
        }
    }
    
    /**
     * Get the current base URL based on provider
     * @returns {string} - Base URL
     */
    getCurrentBaseUrl() {
        return this.provider === 'openrouter' ? this.openRouterBaseUrl : this.baseUrl;
    }
    
    /**
     * Get headers for API requests
     * @returns {object} - Headers object
     */
    getHeaders() {
        const headers = {
            'Content-Type': 'application/json'
        };
        
        if (this.provider === 'openrouter' && this.openRouterApiKey) {
            headers['Authorization'] = `Bearer ${this.openRouterApiKey}`;
        }
        
        return headers;
    }

    /**
     * Check connection to the LLM server
     * @returns {Promise<boolean>}
     */
    async checkConnection() {
        try {
            const baseUrl = this.getCurrentBaseUrl();
            const endpoint = this.provider === 'openrouter' ? '/models' : '/v1/models';
            const response = await fetch(`${baseUrl}${endpoint}`, {
                method: 'GET',
                headers: this.getHeaders(),
                signal: AbortSignal.timeout(5000) // 5 second timeout
            });

            if (response.ok) {
                this.connectionStatus = 'connected';
                // Load available models
                const data = await response.json();
                this.availableModels = data.data || [];
                return true;
            } else {
                this.connectionStatus = 'disconnected';
                return false;
            }
        } catch (error) {
            console.error('Connection error:', error);
            this.connectionStatus = 'disconnected';
            return false;
        }
    }

    /**
     * Get available models from the LLM server
     * @returns {Promise<Array>}
     */
    async getModels() {
        if (this.connectionStatus !== 'connected') {
            await this.checkConnection();
        }
        
        return this.availableModels;
    }

    /**
     * Set the model to use for generation
     * @param {string} modelId 
     */
    setModel(modelId) {
        this.selectedModel = modelId;
        
        // Save selected model for current provider
        if (modelId) {
            const modelKey = `selected_model_${this.provider}`;
            localStorage.setItem(modelKey, modelId);
        }
    }
    
    /**
     * Get the currently selected model
     * @returns {string|null} - Currently selected model ID
     */
    getSelectedModel() {
        return this.selectedModel;
    }

    /**
     * Generate presentation content using chat completions API
     * @param {string} prompt 
     * @param {Array} contexts 
     * @param {string} complexity 
     * @param {number} slideCount 
     * @param {function} onProgress 
     * @returns {Promise<object>}
     */
    async generatePresentation(prompt, contexts = [], complexity = 'standard', slideCount = 10, onProgress, imageLayout = null, useRealImages = false) {
        if (!this.selectedModel) {
            throw new Error('No model selected');
        }

        // Abort any ongoing requests
        if (this.abortController) {
            this.abortController.abort();
        }
        
        this.abortController = new AbortController();
        
        try {
            // Build the system message with instructions
            let systemMessage = `You are an expert presentation creator. 
            Create a professional ${complexity} presentation with exactly ${slideCount} slides based on the prompt.`;
            
            // Add image layout instructions if a layout is selected
            if (imageLayout && imageLayout !== 'none') {
                const imageType = useRealImages ? 'real images from Pexels' : 'image placeholders';
                systemMessage += `\n\nIMPORTANT: You MUST include ${imageType} in EVERY slide. For each slide, you MUST:
                1. Include an "imageLayout" field with the value "${imageLayout}"
                2. Include an "imageDescription" field with a detailed description for finding the right image`;
                
                if (useRealImages) {
                    systemMessage += `\n\nSince real images will be fetched from Pexels, create imageDescription as EXACTLY 2 WORDS that combine:
                    1. The presentation's overall theme: "${prompt.substring(0, 100)}..."
                    2. The specific slide's topic (from the slide title)
                    
                    CRITICAL RULES:
                    - Use EXACTLY 2 words - no more, no less
                    - Choose the most visual/photographable terms
                    - Combine presentation context with slide specifics
                    - Focus on concrete nouns and descriptive terms
                    - First word: usually a descriptor/context
                    - Second word: usually the main subject/noun
                    
                    EXAMPLES showing context + slide title → 2-word query:
                    
                    Presentation: "Digital transformation in healthcare"
                    - Slide: "Patient Care Technology" → "medical technology"
                    - Slide: "Staff Training Programs" → "healthcare training"
                    - Slide: "ROI Analysis" → "healthcare analytics"
                    
                    Presentation: "Startup growth strategies"
                    - Slide: "Building Your Team" → "startup team"
                    - Slide: "Market Expansion" → "global business"
                    - Slide: "Funding Rounds" → "startup investment"
                    
                    Presentation: "Sustainable business practices"
                    - Slide: "Green Operations" → "sustainable business"
                    - Slide: "Employee Engagement" → "team collaboration"
                    - Slide: "Cost Savings" → "green finance"
                    
                    Presentation: "Educational technology"
                    - Slide: "Virtual Classrooms" → "online education"
                    - Slide: "Student Performance" → "student success"
                    - Slide: "Teacher Resources" → "teacher technology"
                    
                    BAD examples (avoid):
                    - "growth" (only 1 word)
                    - "business meeting discussion" (3 words)
                    - "abstract concept" (not photographable)
                    - "slide about" (includes meta language)
                    
                    Format: Return ONLY the 2 words, nothing else`;
                }
                
                let layoutDescription = '';
                switch (imageLayout) {
                    case 'full-width':
                        layoutDescription = 'Image at top, text below (for slides with short text)';
                        break;
                    case 'side-by-side':
                        layoutDescription = 'Image on left, text on right (for balanced content)';
                        break;
                    case 'text-focus':
                        layoutDescription = 'Small image on right, more text space (for text-heavy slides)';
                        break;
                    case 'background':
                        layoutDescription = 'Full background image with text overlay (for impact slides)';
                        break;
                }
                
                systemMessage += `\n\nThe selected layout is "${imageLayout}" - ${layoutDescription}`;
                systemMessage += `\n\nCRITICAL: Every slide object in your response MUST have:
                - "imageLayout": "${imageLayout}"
                - "imageDescription": ${useRealImages ? 'EXACTLY 2 words that combine presentation theme with slide topic' : 'A detailed description of an appropriate image for that slide'}`;
            }
            
            systemMessage += `\n\nFormat your response as a JSON object with the following structure:
            {
                "title": "Presentation Title",
                "slides": [
                    {
                        "title": "Slide 1 Title",
                        "content": ["Bullet point 1", "Bullet point 2"],
                        "notes": "Speaker notes for this slide"${imageLayout && imageLayout !== 'none' ? `,\n                        "imageLayout": "${imageLayout}",\n                        "imageDescription": "Suggested image: [description of relevant image]"` : ''}
                    }
                ]
            }`;
            
            // Build user message with prompt and context
            let userMessage = prompt;
            if (contexts.length > 0) {
                userMessage += "\n\nContext information:\n" + contexts.join("\n\n");
            }
            
            // Create the messages array
            const messages = [
                { role: "system", content: systemMessage },
                { role: "user", content: userMessage }
            ];
            
            // Make the API call with streaming enabled
            const baseUrl = this.getCurrentBaseUrl();
            const endpoint = this.provider === 'openrouter' ? '/chat/completions' : '/v1/chat/completions';
            const response = await fetch(`${baseUrl}${endpoint}`, {
                method: 'POST',
                headers: this.getHeaders(),
                body: JSON.stringify({
                    model: this.selectedModel,
                    messages: messages,
                    stream: true,
                    temperature: 0.7
                }),
                signal: this.abortController.signal
            });
            
            if (!response.ok) {
                const error = await response.text();
                throw new Error(`API error: ${error}`);
            }
            
            // Process the streaming response
            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let content = "";
            let chunkCount = 0;
            let estimatedTotal = slideCount * 200; // rough estimate of token count
            
            while (true) {
                const { done, value } = await reader.read();
                
                if (done) {
                    break;
                }
                
                // Decode the chunk
                const chunk = decoder.decode(value, { stream: true });
                
                // Parse SSE format
                const lines = chunk.split('\n');
                for (const line of lines) {
                    if (line.startsWith('data: ')) {
                        const data = line.slice(6);
                        
                        // Skip [DONE] message
                        if (data === '[DONE]') continue;
                        
                        try {
                            const parsed = JSON.parse(data);
                            if (parsed.choices && parsed.choices[0].delta.content) {
                                content += parsed.choices[0].delta.content;
                                chunkCount++;
                                
                                // Update progress (rough estimate)
                                if (onProgress && typeof onProgress === 'function') {
                                    const progress = Math.min(Math.floor((chunkCount / estimatedTotal) * 100), 99);
                                    onProgress(progress);
                                }
                            }
                        } catch (e) {
                            console.error('Error parsing chunk:', e);
                        }
                    }
                }
            }
            
            // Set final progress to 100%
            if (onProgress && typeof onProgress === 'function') {
                onProgress(100);
            }
            
            // Parse the complete JSON response
            try {
                console.log('Raw LLM response:', content);
                console.log('Image layout selected:', imageLayout);
                
                // Find the JSON object in the text (it might be surrounded by extra text)
                const jsonMatch = content.match(/\{[\s\S]*\}/);
                let presentationData;
                
                if (jsonMatch) {
                    presentationData = JSON.parse(jsonMatch[0]);
                    console.log('Parsed presentation data:', presentationData);
                } else {
                    // Fallback: try to create a structured presentation from the text
                    console.warn('Could not extract JSON from response, generating structured data');
                    
                    // Split content by newlines to find potential slide content
                    const lines = content.split('\n').filter(line => line.trim());
                    
                    // Find potential title (first non-empty line)
                    const title = lines[0] || 'Generated Presentation';
                    
                    // Generate structured slides from content
                    const slides = [];
                    let currentSlide = null;
                    
                    for (let i = 1; i < lines.length; i++) {
                        const line = lines[i].trim();
                        
                        // Skip empty lines
                        if (!line) continue;
                        
                        // Check if this looks like a slide title (shorter line)
                        if (line.length < 50 && !line.startsWith('-') && !line.startsWith('•')) {
                            currentSlide = {
                                title: line,
                                content: [],
                                notes: ''
                            };
                            slides.push(currentSlide);
                        } 
                        // If this looks like a bullet point, add to current slide
                        else if (currentSlide && (line.startsWith('-') || line.startsWith('•'))) {
                            const content = line.substring(1).trim();
                            currentSlide.content.push(content);
                        }
                        // If no slide yet, create first slide
                        else if (!currentSlide) {
                            currentSlide = {
                                title: 'Introduction',
                                content: [line],
                                notes: ''
                            };
                            slides.push(currentSlide);
                        }
                        // Otherwise add as content to current slide
                        else if (currentSlide) {
                            currentSlide.content.push(line);
                        }
                    }
                    
                    // Ensure we have at least some slides
                    if (slides.length === 0) {
                        slides.push({
                            title: 'Content',
                            content: ['Generated from prompt'],
                            notes: ''
                        });
                    }
                    
                    presentationData = {
                        title: title,
                        slides: slides
                    };
                }
                
                // Validate and ensure the presentation data has the required structure
                if (!presentationData.title) {
                    presentationData.title = 'Generated Presentation';
                }
                
                if (!presentationData.slides || !Array.isArray(presentationData.slides) || presentationData.slides.length === 0) {
                    presentationData.slides = [{
                        title: 'Generated Content',
                        content: ['Content generated from your prompt'],
                        notes: ''
                    }];
                }
                
                // Ensure each slide has the required fields
                presentationData.slides = presentationData.slides.map(slide => {
                    return {
                        title: slide.title || 'Slide',
                        content: Array.isArray(slide.content) ? slide.content : [String(slide.content || '')],
                        notes: slide.notes || ''
                    };
                });
                
                return presentationData;
            } catch (error) {
                console.error('Error parsing response:', error);
                // Provide a fallback presentation structure
                return {
                    title: 'Generated Presentation',
                    slides: [
                        {
                            title: 'Generated Content',
                            content: ['Content generated from your prompt'],
                            notes: 'Error occurred during generation'
                        }
                    ]
                };
            }
            
        } catch (error) {
            if (error.name === 'AbortError') {
                throw new Error('Request was cancelled');
            }
            console.error('Generation error:', error);
            throw error;
        } finally {
            this.abortController = null;
        }
    }
    
    /**
     * Generate a single slide based on title and description
     * @param {string} title - Slide title
     * @param {string} description - Slide description (optional)
     * @param {string} complexity - Complexity level ('simple', 'standard', 'detailed')
     * @param {object} context - Additional context for better generation
     * @returns {Promise<object>} - Generated slide data
     */
    async generateSingleSlide(title, description = '', complexity = 'standard', context = {}, imageLayout = null, useRealImages = false) {
        if (!this.selectedModel) {
            throw new Error('No model selected');
        }

        // Abort any ongoing requests
        if (this.abortController) {
            this.abortController.abort();
        }
        
        this.abortController = new AbortController();
        
        try {
            // Build the system message for single slide generation
            let systemMessage = `You are an expert presentation creator. Create a single professional ${complexity} slide that fits seamlessly into an existing presentation.`;
            
            // Add image layout instructions if layouts are selected
            let jsonFormat = `{
    "title": "Slide Title",
    "content": ["Bullet point 1", "Bullet point 2", "Bullet point 3"],
    "notes": "Speaker notes for this slide"`;
            
            if (imageLayout && imageLayout !== 'none') {
                const imageType = useRealImages ? 'real image from Pexels' : 'image placeholder';
                systemMessage += `\n\nFor this slide, you MUST use the "${imageLayout}" layout with a ${imageType}.`;
                
                if (useRealImages) {
                    const presentationContext = context.originalPrompt ? context.originalPrompt.substring(0, 100) : 'business presentation';
                    systemMessage += `\n\nSince a real image will be fetched from Pexels, create imageDescription as EXACTLY 2 WORDS that intelligently combine:
                    1. The overall presentation theme: "${presentationContext}..."
                    2. This specific slide's title: "${title}"
                    
                    RULES FOR 2-WORD QUERIES:
                    - Use EXACTLY 2 words - no more, no less
                    - Blend presentation context with slide specifics
                    - Choose concrete, visual terms
                    - First word: descriptor/context
                    - Second word: main subject/noun
                    
                    CONTEXTUAL EXAMPLES:
                    
                    Presentation: "Digital marketing strategy" + Slide: "Social Media ROI"
                    → "social analytics"
                    
                    Presentation: "Company culture transformation" + Slide: "Remote Work Benefits"
                    → "remote team"
                    
                    Presentation: "Financial planning guide" + Slide: "Investment Portfolio"
                    → "investment portfolio"
                    
                    Presentation: "Product launch plan" + Slide: "Market Research"
                    → "market research"
                    
                    Presentation: "Healthcare innovation" + Slide: "Patient Experience"
                    → "patient care"
                    
                    Presentation: "Education technology" + Slide: "Online Learning"
                    → "online education"
                    
                    CONSIDER ADJACENT SLIDES:
                    ${context.previousSlide ? `Previous slide was about: "${context.previousSlide.title}"` : ''}
                    ${context.nextSlide ? `Next slide will be about: "${context.nextSlide.title}"` : ''}
                    
                    Create a query that bridges the presentation theme with this slide's specific focus.
                    
                    Return ONLY the 2 words`;
                }
                
                let layoutDescription = '';
                switch (imageLayout) {
                    case 'full-width':
                        layoutDescription = 'Image at top, text below (for short text)';
                        break;
                    case 'side-by-side':
                        layoutDescription = 'Image on left, text on right (for balanced content)';
                        break;
                    case 'text-focus':
                        layoutDescription = 'Small image on right, more text space (for text-heavy slides)';
                        break;
                    case 'background':
                        layoutDescription = 'Full background image with text overlay (for impact slides)';
                        break;
                }
                
                systemMessage += `\nLayout: "${imageLayout}" - ${layoutDescription}`;
                
                jsonFormat += `,
    "imageLayout": "${imageLayout}",
    "imageDescription": "${useRealImages ? '[EXACTLY 2 words like "team meeting"]' : 'Suggested image: [description of relevant image]'}"`;
            }
            
            jsonFormat += `
}`;
            
            systemMessage += `\n\nIMPORTANT: Respond with ONLY a valid JSON object in this exact format:
${jsonFormat}

Content guidelines:
- For 'simple' complexity: 2-3 concise bullet points
- For 'standard' complexity: 3-5 informative bullet points  
- For 'detailed' complexity: 5-7 comprehensive bullet points

Context awareness:
- Consider the overall presentation theme and original prompt
- Ensure content flows naturally from the previous slide to the next slide
- Avoid repeating information from adjacent slides
- Maintain consistency with the presentation's tone and style

Make the content substantive, relevant, and professional. Focus on key insights, facts, or actionable information related to the slide title.`;
            
            // Build user message with title, description, and context
            let userMessage = `Title: ${title}`;
            if (description) {
                userMessage += `\nDescription: ${description}`;
            }
            
            // Add context if provided
            if (context.originalPrompt) {
                userMessage += `\n\nOriginal presentation prompt: ${context.originalPrompt}`;
            }
            
            if (context.presentationTitle) {
                userMessage += `\nPresentation title: ${context.presentationTitle}`;
            }
            
            if (context.position !== undefined) {
                userMessage += `\n\nThis slide will be inserted at position ${context.position + 1} in the presentation.`;
            }
            
            if (context.previousSlide) {
                userMessage += `\n\nPrevious slide (${context.previousSlide.title}):\n`;
                if (Array.isArray(context.previousSlide.content)) {
                    context.previousSlide.content.forEach(item => {
                        userMessage += `- ${item}\n`;
                    });
                } else if (context.previousSlide.content) {
                    userMessage += context.previousSlide.content;
                }
            }
            
            if (context.nextSlide) {
                userMessage += `\n\nNext slide (${context.nextSlide.title}):\n`;
                if (Array.isArray(context.nextSlide.content)) {
                    context.nextSlide.content.forEach(item => {
                        userMessage += `- ${item}\n`;
                    });
                } else if (context.nextSlide.content) {
                    userMessage += context.nextSlide.content;
                }
            }
            
            // Create the messages array
            const messages = [
                { role: "system", content: systemMessage },
                { role: "user", content: userMessage }
            ];
            
            // Make the API call
            const baseUrl = this.getCurrentBaseUrl();
            const endpoint = this.provider === 'openrouter' ? '/chat/completions' : '/v1/chat/completions';
            const response = await fetch(`${baseUrl}${endpoint}`, {
                method: 'POST',
                headers: this.getHeaders(),
                body: JSON.stringify({
                    model: this.selectedModel,
                    messages: messages,
                    stream: false,
                    temperature: 0.7
                }),
                signal: this.abortController.signal
            });
            
            if (!response.ok) {
                const error = await response.text();
                throw new Error(`API error: ${error}`);
            }
            
            const data = await response.json();
            
            // Extract the content from the response
            if (data.choices && data.choices[0] && data.choices[0].message) {
                const content = data.choices[0].message.content;
                console.log('Raw LLM response for single slide:', content);
                
                try {
                    // Try to find JSON in the response (it might be wrapped in other text)
                    const jsonMatch = content.match(/\{[\s\S]*?\}/);
                    let slideData;
                    
                    if (jsonMatch) {
                        slideData = JSON.parse(jsonMatch[0]);
                    } else {
                        // Try parsing the entire content as JSON
                        slideData = JSON.parse(content);
                    }
                    
                    // Validate and normalize the slide data structure
                    const normalizedSlide = {
                        title: slideData.title || title,
                        content: Array.isArray(slideData.content) ? slideData.content : 
                                slideData.content ? [slideData.content] : 
                                description ? [description] : ['Generated content based on: ' + title],
                        notes: slideData.notes || ''
                    };
                    
                    console.log('Parsed slide data:', normalizedSlide);
                    return normalizedSlide;
                    
                } catch (parseError) {
                    console.error('Error parsing slide data:', parseError);
                    console.log('Content that failed to parse:', content);
                    
                    // Enhanced fallback: try to extract meaningful content from the response
                    const lines = content.split('\n').filter(line => line.trim());
                    const fallbackContent = [];
                    
                    // Look for bullet points or numbered lists
                    for (const line of lines) {
                        const trimmedLine = line.trim();
                        if (trimmedLine.startsWith('-') || trimmedLine.startsWith('•') || 
                            trimmedLine.startsWith('*') || /^\d+\./.test(trimmedLine)) {
                            // Clean up the bullet point
                            fallbackContent.push(trimmedLine.replace(/^[-•*]\s*/, '').replace(/^\d+\.\s*/, ''));
                        } else if (trimmedLine.length > 10 && !trimmedLine.includes('{') && !trimmedLine.includes('}')) {
                            // Add lines that look like content (not JSON fragments)
                            fallbackContent.push(trimmedLine);
                        }
                    }
                    
                    // If we found some content, use it; otherwise use description or title-based content
                    const finalContent = fallbackContent.length > 0 ? fallbackContent :
                                       description ? [description] : 
                                       [`Key points about ${title}`, 'Generated content based on the slide title'];
                    
                    return {
                        title: title,
                        content: finalContent,
                        notes: ''
                    };
                }
            } else {
                throw new Error('Invalid response format');
            }
            
        } catch (error) {
            console.error('Single slide generation error:', error);
            throw error;
        } finally {
            this.abortController = null;
        }
    }
    
    /**
     * Cancel an ongoing generation request
     */
    cancelGeneration() {
        if (this.abortController) {
            this.abortController.abort();
            this.abortController = null;
        }
    }
}

// Create and export a singleton instance
const apiClient = new ApiClient();
