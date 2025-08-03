/**
 * Pexels API Client for fetching real images
 */
class PexelsClient {
    constructor() {
        this.baseUrl = 'https://api.pexels.com/v1';
        this.apiKey = '';
        this.imageCache = new Map();
        this.failedSearchCache = new Map(); // Cache failed searches
        this.cacheExpiry = 900000; // 15 minutes cache expiry
        this.rateLimitInfo = {
            limit: 200,
            remaining: 200,
            reset: null
        };
        
        // Rate limiting queue system
        this.requestQueue = [];
        this.processing = false;
        this.lastRequestTime = 0;
        this.minDelay = 200; // 200ms between requests (5 requests/second max)
        this.requestsInWindow = [];
        this.windowSize = 3600000; // 1 hour in ms
        this.maxRequestsPerHour = 180; // Leave buffer from 200 limit
    }

    /**
     * Set the Pexels API key
     * @param {string} apiKey - Pexels API key
     */
    setApiKey(apiKey) {
        this.apiKey = apiKey;
        localStorage.setItem('pexels_api_key', apiKey);
    }

    /**
     * Load saved API key from localStorage
     */
    loadSavedApiKey() {
        const savedKey = localStorage.getItem('pexels_api_key');
        if (savedKey) {
            this.apiKey = savedKey;
        }
        return savedKey;
    }

    /**
     * Get headers for API requests
     * @returns {object} - Headers object
     */
    getHeaders() {
        return {
            'Authorization': this.apiKey
        };
    }

    /**
     * Check if API key is configured
     * @returns {boolean}
     */
    hasApiKey() {
        return this.apiKey && this.apiKey.length > 0;
    }
    
    /**
     * Get shared concept mappings for consistent search optimization
     * @returns {object} - Concept mappings
     */
    getConceptMappings() {
        return {
            // Business concepts
            'teamwork': ['team collaboration', 'business team', 'office teamwork', 'colleagues working'],
            'innovation': ['creative innovation', 'technology innovation', 'business innovation', 'lightbulb idea'],
            'leadership': ['business leader', 'executive', 'CEO', 'management'],
            'success': ['business success', 'achievement', 'victory celebration', 'winning'],
            'growth': ['business growth', 'growth chart', 'financial growth', 'progress arrow'],
            'strategy': ['business strategy', 'chess', 'planning', 'roadmap'],
            'productivity': ['productive office', 'efficient work', 'busy workplace', 'working hard'],
            'partnership': ['business handshake', 'partnership', 'business deal', 'cooperation'],
            
            // Technical concepts
            'technology': ['computer technology', 'digital technology', 'tech innovation', 'laptop'],
            'data': ['data analytics', 'big data', 'data visualization', 'statistics chart'],
            'security': ['cyber security', 'data protection', 'security lock', 'shield protection'],
            'ai': ['artificial intelligence', 'AI robot', 'machine learning', 'futuristic tech'],
            'cloud': ['cloud computing', 'cloud storage', 'server room', 'data center'],
            'development': ['software development', 'coding', 'programming', 'developer'],
            'digital': ['digital transformation', 'digital innovation', 'digital business'],
            'software': ['software technology', 'software solution', 'software platform'],
            
            // Communication concepts
            'communication': ['business meeting', 'discussion', 'presentation', 'conference'],
            'marketing': ['digital marketing', 'advertising', 'social media', 'promotion'],
            'customer': ['customer service', 'client meeting', 'customer satisfaction', 'happy customer'],
            'sales': ['sales team', 'sales chart', 'retail', 'commerce'],
            'presentation': ['business presentation', 'presenting', 'conference presentation'],
            'meeting': ['team meeting', 'business meeting', 'office meeting'],
            
            // Finance concepts
            'finance': ['financial', 'money', 'investment', 'banking'],
            'budget': ['budget planning', 'financial planning', 'calculator', 'accounting'],
            'profit': ['profit growth', 'revenue', 'earnings', 'financial success'],
            'investment': ['stock market', 'investment portfolio', 'trading', 'investor'],
            'revenue': ['business revenue', 'income growth', 'financial earnings'],
            'economy': ['global economy', 'economic growth', 'market economy'],
            
            // Healthcare concepts
            'healthcare': ['medical care', 'healthcare system', 'hospital healthcare'],
            'health': ['healthcare', 'wellness', 'medical', 'healthy lifestyle'],
            'medical': ['medical technology', 'medical equipment', 'medical professionals'],
            'patient': ['patient care', 'patient experience', 'medical patient'],
            'wellness': ['health wellness', 'wellness program', 'healthy living'],
            'hospital': ['hospital healthcare', 'medical center', 'healthcare facility'],
            
            // Education concepts
            'education': ['education', 'learning', 'classroom', 'students studying'],
            'learning': ['student learning', 'online learning', 'education learning'],
            'training': ['professional training', 'employee training', 'skill training'],
            'teaching': ['teacher classroom', 'teaching students', 'education teaching'],
            'student': ['students studying', 'student success', 'student learning'],
            
            // Other domains
            'nature': ['nature landscape', 'environment', 'green nature', 'outdoor'],
            'environment': ['green environment', 'sustainable environment', 'natural environment'],
            'sustainability': ['sustainable business', 'green sustainability', 'eco friendly'],
            'global': ['global business', 'world map', 'international', 'globe'],
            'future': ['future technology', 'futuristic', 'tomorrow', 'next generation'],
            'quality': ['high quality', 'premium', 'excellence', 'professional'],
            'startup': ['startup business', 'startup team', 'startup office'],
            'remote': ['remote work', 'work from home', 'remote team']
        };
    }
    
    /**
     * Generate normalized cache key
     * @param {string} query - Search query
     * @param {object} options - Search options
     * @returns {string} - Cache key
     */
    getCacheKey(query, options = {}) {
        // Normalize query for better cache hits
        const normalizedQuery = query.toLowerCase().split(' ').sort().join(' ');
        const optionString = JSON.stringify({
            orientation: options.orientation,
            page: options.page || 1,
            per_page: options.per_page || 15
        });
        return `search-${normalizedQuery}-${optionString}`;
    }
    
    /**
     * Check if cache entry is still valid
     * @param {object} cacheEntry - Cache entry with timestamp
     * @returns {boolean} - Whether cache is valid
     */
    isCacheValid(cacheEntry) {
        if (!cacheEntry || !cacheEntry.timestamp) return false;
        return Date.now() - cacheEntry.timestamp < this.cacheExpiry;
    }
    
    /**
     * Set cache with timestamp
     * @param {string} key - Cache key
     * @param {any} data - Data to cache
     * @param {boolean} isFailure - Whether this is a failed search
     */
    setCache(key, data, isFailure = false) {
        const cache = isFailure ? this.failedSearchCache : this.imageCache;
        cache.set(key, {
            data: data,
            timestamp: Date.now()
        });
        
        // Clean old cache entries if too large
        if (cache.size > 100) {
            // Remove oldest entries
            const entries = Array.from(cache.entries());
            entries.sort((a, b) => a[1].timestamp - b[1].timestamp);
            // Remove oldest 20 entries
            for (let i = 0; i < 20; i++) {
                cache.delete(entries[i][0]);
            }
        }
    }
    
    /**
     * Get from cache if valid
     * @param {string} key - Cache key
     * @param {boolean} checkFailures - Whether to check failed search cache
     * @returns {any|null} - Cached data or null
     */
    getFromCache(key, checkFailures = false) {
        // Check success cache first
        const cached = this.imageCache.get(key);
        if (cached && this.isCacheValid(cached)) {
            return cached.data;
        }
        
        // Check failure cache if requested
        if (checkFailures) {
            const failed = this.failedSearchCache.get(key);
            if (failed && this.isCacheValid(failed)) {
                return null; // Indicate cached failure
            }
        }
        
        return undefined; // Not in cache
    }

    /**
     * Queue a request to avoid rate limiting
     * @param {Function} requestFunction - Function to execute
     * @returns {Promise<any>} - Result of the request
     */
    async queueRequest(requestFunction) {
        return new Promise((resolve, reject) => {
            this.requestQueue.push({ requestFunction, resolve, reject });
            this.processQueue();
        });
    }
    
    /**
     * Process the request queue with rate limiting
     */
    async processQueue() {
        if (this.processing || this.requestQueue.length === 0) return;
        
        this.processing = true;
        const now = Date.now();
        const timeSinceLastRequest = now - this.lastRequestTime;
        
        // Ensure minimum delay between requests
        if (timeSinceLastRequest < this.minDelay) {
            await new Promise(resolve => 
                setTimeout(resolve, this.minDelay - timeSinceLastRequest)
            );
        }
        
        // Clean old requests from window
        this.cleanOldRequests();
        
        // Check hourly limit
        if (this.requestsInWindow.length >= this.maxRequestsPerHour) {
            console.warn(`Rate limit approaching: ${this.requestsInWindow.length}/${this.maxRequestsPerHour} requests in current hour`);
            // Wait until oldest request expires
            const oldestRequest = this.requestsInWindow[0];
            const waitTime = this.windowSize - (now - oldestRequest);
            if (waitTime > 0) {
                console.log(`Waiting ${Math.ceil(waitTime / 1000)}s for rate limit window...`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                this.cleanOldRequests();
            }
        }
        
        // Process next request
        const { requestFunction, resolve, reject } = this.requestQueue.shift();
        this.lastRequestTime = Date.now();
        this.requestsInWindow.push(this.lastRequestTime);
        
        try {
            const result = await requestFunction();
            resolve(result);
        } catch (error) {
            // Handle rate limit errors specifically
            if (error.message && error.message.includes('429')) {
                console.error('Rate limit exceeded, queuing for retry...');
                // Re-queue the request
                this.requestQueue.unshift({ requestFunction, resolve, reject });
                // Wait longer before retrying
                await new Promise(resolve => setTimeout(resolve, 60000)); // Wait 1 minute
            } else {
                reject(error);
            }
        }
        
        this.processing = false;
        // Process next in queue
        setTimeout(() => this.processQueue(), 0);
    }
    
    /**
     * Clean old requests from the tracking window
     */
    cleanOldRequests() {
        const now = Date.now();
        this.requestsInWindow = this.requestsInWindow.filter(
            timestamp => now - timestamp < this.windowSize
        );
    }
    
    /**
     * Search photos by query
     * @param {string} query - Search query
     * @param {object} options - Search options
     * @returns {Promise<object>} - Search results
     */
    async searchPhotos(query, options = {}) {
        const {
            page = 1,
            per_page = 15,
            orientation = 'landscape',
            size = 'large',
            color = null,
            locale = 'en-US'
        } = options;

        // Check cache first
        const cacheKey = this.getCacheKey(query, options);
        const cachedResult = this.getFromCache(cacheKey, true);
        
        if (cachedResult !== undefined) {
            if (cachedResult === null) {
                console.log(`Cached failure for query: "${query}" - skipping retry`);
                return { photos: [] };
            } else {
                console.log(`Cache hit for query: "${query}"`);
                return cachedResult;
            }
        }
        
        // Use rate-limited queue for API requests
        return await this.queueRequest(async () => {
            let url = `${this.baseUrl}/search?query=${encodeURIComponent(query)}`;
            url += `&page=${page}`;
            url += `&per_page=${per_page}`;
            url += `&orientation=${orientation}`;
            url += `&size=${size}`;
            url += `&locale=${locale}`;
            
            if (color) {
                url += `&color=${color}`;
            }

            const response = await fetch(url, {
                method: 'GET',
                headers: this.getHeaders()
            });

            // Update rate limit info
            this.updateRateLimitInfo(response.headers);

            if (!response.ok) {
                if (response.status === 401) {
                    throw new Error('Invalid Pexels API key. Please check your settings.');
                } else if (response.status === 429) {
                    throw new Error('429: Pexels API rate limit exceeded.');
                }
                throw new Error(`Pexels API error: ${response.status}`);
            }

            const data = await response.json();
            
            // Cache the results (success or empty)
            if (data.photos && data.photos.length > 0) {
                this.setCache(cacheKey, data, false);
            } else {
                // Cache empty results as failures to avoid retrying
                this.setCache(cacheKey, data, true);
            }

            return data;
        });
    }

    /**
     * Get image for slide based on context
     * @param {string} context - Slide context/description
     * @param {string} layout - Image layout type
     * @param {object} themeColors - Optional theme colors for filtering
     * @returns {Promise<object|null>} - Image data or null
     */
    async getImageForSlide(context, layout = 'landscape', themeColors = null) {
        try {
            // Determine orientation based on layout
            let orientation = 'landscape';
            if (layout === 'background' || layout === 'full-width') {
                orientation = 'landscape';
            } else if (layout === 'side-by-side' || layout === 'text-focus') {
                orientation = 'square';
            }
            
            console.log(`\n=== Searching for image ===`);
            console.log(`Original context: "${context}"`);
            console.log(`Layout: ${layout}, Orientation: ${orientation}`);
            
            // Log current rate limit status
            const status = this.getRateLimitStatus();
            if (status.requestsInCurrentHour > status.maxRequestsPerHour * 0.8) {
                console.warn(`⚠️ Approaching rate limit: ${status.requestsInCurrentHour}/${status.maxRequestsPerHour} requests this hour`);
            }
            
            // Check if this is a pre-optimized query (1 or 2 words)
            const wordCount = context.trim().split(/\s+/).length;
            const isPreOptimized = wordCount <= 2;
            console.log(`Query type: ${wordCount}-word${isPreOptimized ? ' (pre-optimized)' : ' (needs optimization)'}`);
            
            // Strategy 1: Primary search with optimized query and variations
            const searchResults = await this.primarySearchStrategy(context, orientation, themeColors, isPreOptimized);
            if (searchResults) {
                console.log(`✓ Found image using primary search strategy`);
                return searchResults;
            }
            
            // Strategy 2: Concept-based search
            const conceptResults = await this.conceptSearchStrategy(context, orientation);
            if (conceptResults) {
                console.log(`✓ Found image using concept search strategy`);
                return conceptResults;
            }
            
            // Strategy 3: Broader search with main keywords
            const broadResults = await this.broadSearchStrategy(context, orientation);
            if (broadResults) {
                console.log(`✓ Found image using broad search strategy`);
                return broadResults;
            }
            
            // Strategy 4: Curated fallback
            const curatedResults = await this.curatedFallbackStrategy(context);
            if (curatedResults) {
                console.log(`✓ Found image using curated fallback`);
                return curatedResults;
            }

            console.log('✗ No suitable images found');
            return null;
        } catch (error) {
            console.error('Error getting image for slide:', error);
            return null;
        }
    }
    
    /**
     * Get curated photos as fallback
     * @param {object} options - API options
     * @returns {Promise<object>} - Curated photos response
     */
    async getCuratedPhotos(options = {}) {
        const { page = 1, per_page = 15 } = options;
        
        // Check cache first
        const cacheKey = `curated-${page}-${per_page}`;
        const cachedResult = this.getFromCache(cacheKey, false);
        if (cachedResult !== undefined) {
            console.log(`Cache hit for curated photos`);
            return cachedResult;
        }
        
        // Use rate-limited queue for API requests
        return await this.queueRequest(async () => {
            const url = `${this.baseUrl}/curated?page=${page}&per_page=${per_page}`;
            
            const response = await fetch(url, {
                method: 'GET',
                headers: this.getHeaders()
            });

            if (!response.ok) {
                if (response.status === 429) {
                    throw new Error('429: Pexels API rate limit exceeded.');
                }
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            
            // Cache the results
            this.setCache(cacheKey, data, false);
            
            return data;
        });
    }
    
    /**
     * Rank photos by relevance to context
     * @param {Array} photos - Array of photos from search
     * @param {string} originalContext - Original search context
     * @param {object} themeColors - Theme colors for matching
     * @returns {Array} - Ranked photos
     */
    rankPhotosByRelevance(photos, originalContext, themeColors) {
        const contextWords = originalContext.toLowerCase().split(/\s+/)
            .filter(word => word.length > 2);
        
        // Extract key concepts for better matching
        const concepts = this.extractConcepts(originalContext);
        const visualKeywords = this.extractVisualKeywords(this.cleanQuery(originalContext));
        
        return photos.map(photo => {
            let score = 0;
            
            // 1. Alt text relevance (highest weight)
            if (photo.alt) {
                const altLower = photo.alt.toLowerCase();
                
                // Exact word matches
                contextWords.forEach(word => {
                    if (word.length > 3) {
                        if (altLower.includes(word)) {
                            score += 3;
                        }
                        // Partial matches
                        else if (altLower.includes(word.slice(0, -1)) || altLower.includes(word + 's')) {
                            score += 1.5;
                        }
                    }
                });
                
                // Concept matches
                concepts.forEach(concept => {
                    if (altLower.includes(concept.toLowerCase())) {
                        score += 4;
                    }
                });
                
                // Visual keyword matches
                visualKeywords.forEach(keyword => {
                    if (altLower.includes(keyword.toLowerCase())) {
                        score += 2;
                    }
                });
            }
            
            // 2. Photo quality indicators
            if (photo.width >= 3000 && photo.height >= 2000) {
                score += 1; // High resolution
            }
            
            if (photo.alt && photo.alt.length > 20) {
                score += 0.5; // Well-described photo
            }
            
            // 3. Professional/business context bonus
            if (photo.alt) {
                const businessTerms = ['business', 'office', 'professional', 'corporate', 'meeting', 'team'];
                businessTerms.forEach(term => {
                    if (photo.alt.toLowerCase().includes(term)) {
                        score += 0.5;
                    }
                });
            }
            
            // 4. Photographer relevance (some photographers specialize)
            if (photo.photographer) {
                const photographerLower = photo.photographer.toLowerCase();
                if (contextWords.some(word => photographerLower.includes(word))) {
                    score += 0.5;
                }
            }
            
            // 5. Color relevance (if theme colors provided)
            if (themeColors && photo.avg_color) {
                score += this.calculateColorRelevance(photo.avg_color, originalContext, themeColors);
            }
            
            // 6. Popularity indicator
            if (photo.liked === true) {
                score += 0.3;
            }
            
            return { ...photo, relevanceScore: score };
        })
        .sort((a, b) => b.relevanceScore - a.relevanceScore)
        .filter(photo => photo.relevanceScore > 0);
    }
    
    /**
     * Calculate color relevance based on context and theme
     */
    calculateColorRelevance(avgColor, context, themeColors) {
        let score = 0;
        const contextLower = context.toLowerCase();
        
        // Color associations for different contexts
        const colorContextMap = {
            'business': ['#0066CC', '#003366', '#333333', '#666666'],
            'professional': ['#2C3E50', '#34495E', '#1A1A1A', '#4A4A4A'],
            'technology': ['#0099FF', '#00CCFF', '#333333', '#0066CC'],
            'nature': ['#228B22', '#008000', '#90EE90', '#32CD32'],
            'creative': ['#FF6B6B', '#4ECDC4', '#FFE66D', '#A8E6CF'],
            'finance': ['#006400', '#008000', '#FFD700', '#0066CC'],
            'health': ['#32CD32', '#00CED1', '#FFFFFF', '#87CEEB']
        };
        
        // Check if context matches any color associations
        for (const [contextKey, colors] of Object.entries(colorContextMap)) {
            if (contextLower.includes(contextKey)) {
                colors.forEach(color => {
                    if (this.colorsAreSimilar(avgColor, color)) {
                        score += 1;
                    }
                });
            }
        }
        
        // Check theme color matching
        if (themeColors) {
            if (this.colorsAreSimilar(avgColor, '#' + themeColors.primaryColor)) {
                score += 0.5;
            }
            if (this.colorsAreSimilar(avgColor, '#' + themeColors.secondaryColor)) {
                score += 0.3;
            }
        }
        
        return score;
    }
    
    /**
     * Check if two colors are similar
     */
    colorsAreSimilar(color1, color2, threshold = 50) {
        // Convert hex to RGB
        const rgb1 = this.hexToRgb(color1);
        const rgb2 = this.hexToRgb(color2);
        
        if (!rgb1 || !rgb2) return false;
        
        // Calculate color distance
        const distance = Math.sqrt(
            Math.pow(rgb1.r - rgb2.r, 2) +
            Math.pow(rgb1.g - rgb2.g, 2) +
            Math.pow(rgb1.b - rgb2.b, 2)
        );
        
        return distance < threshold;
    }
    
    /**
     * Convert hex color to RGB
     */
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
        } : null;
    }

    /**
     * Format photo data for use in presentation
     * @param {object} photo - Photo object from Pexels API
     * @returns {object} - Formatted photo data
     */
    formatPhotoData(photo) {
        return {
            id: photo.id,
            url: photo.src.large2x || photo.src.large,
            thumbnail: photo.src.medium,
            original: photo.src.original,
            photographer: photo.photographer,
            photographerUrl: photo.photographer_url,
            avgColor: photo.avg_color,
            alt: photo.alt || '',
            width: photo.width,
            height: photo.height,
            attribution: {
                text: `Photo by ${photo.photographer} on Pexels`,
                html: `Photo by <a href="${photo.photographer_url}?utm_source=pexels&utm_medium=referral">${photo.photographer}</a> on <a href="https://www.pexels.com?utm_source=pexels&utm_medium=referral">Pexels</a>`
            }
        };
    }

    /**
     * Optimize search query for better Pexels results
     * @param {string} description - Original image description
     * @returns {string} - Optimized search query
     */
    optimizeSearchQuery(description) {
        // Check if this is already a pre-optimized query (1 or 2 words)
        const words = description.trim().split(/\s+/);
        if (words.length <= 2) {
            console.log(`Using pre-optimized query as-is: "${description}"`);
            return description.trim();
        }
        
        // Remove presentation-specific phrases
        const presentationPhrases = [
            'slide about', 'showing', 'discussing', 'presenting', 'illustrating',
            'demonstrating', 'explaining', 'slide for', 'image of', 'picture of',
            'visual representation', 'graphic showing', 'detailed search query:',
            'suggested image:', 'relevant image', 'for the', 'that shows', 'which displays'
        ];
        
        let query = description.toLowerCase();
        presentationPhrases.forEach(phrase => {
            query = query.replace(new RegExp(phrase, 'gi'), '');
        });
        
        // Clean and extract key visual terms
        const cleanedQuery = this.cleanQuery(query);
        const visualKeywords = this.extractVisualKeywords(cleanedQuery);
        
        // Generate search variations and pick the best one
        const variations = this.generateSearchVariations(visualKeywords);
        
        // Return the primary variation (most relevant)
        return variations[0] || cleanedQuery;
    }
    
    /**
     * Clean query by removing stop words
     * @param {string} query - Query to clean
     * @returns {string} - Cleaned query
     */
    cleanQuery(query) {
        const stopWords = new Set([
            'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for',
            'of', 'with', 'by', 'from', 'about', 'as', 'is', 'was', 'are', 'were',
            'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will',
            'would', 'could', 'should', 'may', 'might', 'must', 'can', 'shall',
            'very', 'really', 'quite', 'just', 'even'
        ]);
        
        return query
            .toLowerCase()
            .split(/\s+/)
            .filter(word => word.length > 2 && !stopWords.has(word))
            .join(' ')
            .trim();
    }
    
    /**
     * Generate search variations for better results
     * @param {Array<string>} keywords - Visual keywords
     * @returns {Array<string>} - Search variations
     */
    generateSearchVariations(keywords) {
        const variations = [];
        
        // Primary variation - top keywords
        if (keywords.length > 0) {
            variations.push(keywords.slice(0, 3).join(' '));
        }
        
        // Add industry prefixes for business context
        const industryPrefixes = ['business', 'professional', 'modern', 'office'];
        const primaryKeyword = keywords[0] || '';
        
        industryPrefixes.forEach(prefix => {
            if (primaryKeyword && !primaryKeyword.includes(prefix)) {
                variations.push(`${prefix} ${primaryKeyword}`);
            }
        });
        
        // Add singular/plural variations
        if (keywords.length > 0) {
            const firstKeyword = keywords[0];
            if (firstKeyword.endsWith('s') && firstKeyword.length > 3) {
                variations.push(firstKeyword.slice(0, -1) + ' ' + keywords.slice(1, 3).join(' '));
            } else if (!firstKeyword.endsWith('s')) {
                variations.push(firstKeyword + 's ' + keywords.slice(1, 3).join(' '));
            }
        }
        
        // Remove duplicates and limit
        return [...new Set(variations)].filter(v => v.length > 0).slice(0, 5);
    }
    
    /**
     * Extract visual keywords that work well with stock photos
     * @param {string} text - Text to extract keywords from
     * @returns {Array<string>} - Array of visual keywords
     */
    extractVisualKeywords(text) {
        // Use shared concept mappings from getConceptMappings()
        const conceptMappings = this.getConceptMappings();
        
        // Split and clean words
        const words = text.toLowerCase()
            .replace(/[^\w\s]/g, ' ')
            .split(/\s+/)
            .filter(word => word.length > 2);
        
        // Prioritize concrete nouns and visual terms
        const priorityTerms = [];
        const secondaryTerms = [];
        const mappedConcepts = new Set();
        
        // First pass: Check for concept mappings
        words.forEach(word => {
            for (const [concept, mappings] of Object.entries(conceptMappings)) {
                if (word === concept || word.includes(concept)) {
                    // Add the best mapping for this concept
                    if (!mappedConcepts.has(concept)) {
                        priorityTerms.push(mappings[0]);
                        mappedConcepts.add(concept);
                    }
                    return;
                }
            }
        });
        
        // Second pass: Check for concrete visual terms
        const visualTerms = [
            'office', 'business', 'people', 'person', 'team', 'meeting', 
            'desk', 'computer', 'laptop', 'phone', 'tablet', 'screen',
            'chart', 'graph', 'diagram', 'presentation', 'whiteboard',
            'money', 'dollar', 'finance', 'calculator', 'document',
            'building', 'city', 'skyline', 'workspace', 'workplace',
            'handshake', 'collaboration', 'discussion', 'conference'
        ];
        
        words.forEach(word => {
            if (!mappedConcepts.has(word)) {
                if (visualTerms.includes(word)) {
                    priorityTerms.push(word);
                } else if (word.length > 3) {
                    secondaryTerms.push(word);
                }
            }
        });
        
        // Combine and deduplicate
        const allTerms = [...new Set([...priorityTerms, ...secondaryTerms])];
        
        // Return top terms, prioritizing mapped concepts and visual terms
        return allTerms.slice(0, 6);
    }
    
    /**
     * Create fallback search queries
     * @param {string} originalQuery - Original optimized query
     * @returns {Array<string>} - Array of fallback queries
     */
    createFallbackQueries(originalQuery) {
        const words = originalQuery.split(' ');
        const fallbacks = [];
        
        // First fallback: Remove adjectives, keep nouns
        if (words.length > 2) {
            fallbacks.push(words.slice(-2).join(' '));
        }
        
        // Second fallback: Category-based
        const categories = {
            'business': ['team', 'office', 'meeting', 'corporate', 'professional'],
            'technology': ['computer', 'laptop', 'digital', 'tech', 'innovation'],
            'finance': ['money', 'chart', 'graph', 'investment', 'financial'],
            'nature': ['landscape', 'outdoor', 'environment', 'natural'],
            'abstract': ['background', 'texture', 'pattern', 'abstract']
        };
        
        for (const [category, keywords] of Object.entries(categories)) {
            if (words.some(word => keywords.some(kw => word.includes(kw)))) {
                fallbacks.push(category);
                break;
            }
        }
        
        // Default fallback
        fallbacks.push('business background');
        
        return [...new Set(fallbacks)];
    }

    /**
     * Update rate limit information from response headers
     * @param {Headers} headers - Response headers
     */
    updateRateLimitInfo(headers) {
        const limit = headers.get('X-Ratelimit-Limit');
        const remaining = headers.get('X-Ratelimit-Remaining');
        const reset = headers.get('X-Ratelimit-Reset');

        if (limit) this.rateLimitInfo.limit = parseInt(limit);
        if (remaining) {
            this.rateLimitInfo.remaining = parseInt(remaining);
            // Log warning if approaching limit
            if (this.rateLimitInfo.remaining < 20) {
                console.warn(`⚠️ Pexels API rate limit warning: Only ${this.rateLimitInfo.remaining} requests remaining!`);
            }
        }
        if (reset) this.rateLimitInfo.reset = new Date(parseInt(reset) * 1000);
        
        // Log current rate limit status for debugging
        console.log(`Rate limit: ${this.rateLimitInfo.remaining}/${this.rateLimitInfo.limit} remaining`);
    }

    /**
     * Get current rate limit status
     * @returns {object} - Rate limit information
     */
    getRateLimitStatus() {
        this.cleanOldRequests();
        return {
            ...this.rateLimitInfo,
            percentageUsed: ((this.rateLimitInfo.limit - this.rateLimitInfo.remaining) / this.rateLimitInfo.limit) * 100,
            requestsInCurrentHour: this.requestsInWindow.length,
            maxRequestsPerHour: this.maxRequestsPerHour,
            queueLength: this.requestQueue.length
        };
    }

    /**
     * Preload image to ensure it's available
     * @param {string} imageUrl - Image URL to preload
     * @returns {Promise<void>}
     */
    async preloadImage(imageUrl) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => resolve();
            img.onerror = () => reject(new Error('Failed to preload image'));
            img.src = imageUrl;
        });
    }

    /**
     * Primary search strategy with optimized queries and variations
     */
    async primarySearchStrategy(context, orientation, themeColors, isPreOptimized = false) {
        let allPhotos = [];
        let variations = [];
        
        if (isPreOptimized) {
            // For pre-optimized queries, use directly first
            console.log(`Primary search - Direct query: "${context}"`);
            console.log(`Strategy: Direct search with pre-optimized ${context.split(' ').length}-word query`);
            
            // Try the exact query first
            const directResults = await this.searchPhotos(context, {
                per_page: 15,
                orientation: orientation
            });
            
            if (directResults.photos && directResults.photos.length > 0) {
                allPhotos = directResults.photos;
                console.log(`Found ${directResults.photos.length} images with direct query`);
            } else {
                console.log(`No results with direct query: "${context}"`);
            }
            
            // Generate minimal variations for pre-optimized queries
            const words = context.trim().split(/\s+/);
            if (words.length === 2) {
                // Try swapped order
                variations.push(`${words[1]} ${words[0]}`);
                // Try each word separately as fallback (prioritize the more specific word)
                const [word1, word2] = words;
                // Determine which word is more specific/visual
                const moreSpecificWord = this.selectMoreSpecificWord(word1, word2);
                const lessSpecificWord = word1 === moreSpecificWord ? word2 : word1;
                variations.push(moreSpecificWord);
                variations.push(lessSpecificWord);
            } else if (words.length === 1) {
                // For single words, try with context-appropriate prefixes
                const word = words[0];
                const prefixes = this.getContextualPrefixes(word);
                prefixes.forEach(prefix => {
                    variations.push(`${prefix} ${word}`);
                });
            }
        } else {
            // For non-optimized queries, use existing optimization
            const optimizedQuery = this.optimizeSearchQuery(context);
            const keywords = this.extractVisualKeywords(this.cleanQuery(context));
            variations = this.generateSearchVariations(keywords);
            
            console.log(`Primary search - Optimized: "${optimizedQuery}"`);
            
            // Try main optimized query first
            const mainResults = await this.searchPhotos(optimizedQuery, {
                per_page: 15,
                orientation: orientation
            });
            
            if (mainResults.photos) {
                allPhotos = mainResults.photos;
            }
        }
        
        console.log(`Variations: ${variations.join(', ')}`);
        
        // Try variations if needed
        if (allPhotos.length < 10 && variations.length > 0) {
            // Try 2-word variations first
            const twoWordVariations = variations.filter(v => v.split(' ').length === 2);
            const oneWordVariations = variations.filter(v => v.split(' ').length === 1);
            
            // Try 2-word variations
            for (const variation of twoWordVariations) {
                if (variation !== context && allPhotos.length < 10) {
                    const varResults = await this.searchPhotos(variation, {
                        per_page: 10,
                        orientation: orientation
                    });
                    if (varResults.photos && varResults.photos.length > 0) {
                        allPhotos = allPhotos.concat(varResults.photos);
                    }
                }
            }
            
            // If still not enough results, try 1-word fallbacks
            if (allPhotos.length < 5) {
                for (const variation of oneWordVariations) {
                    if (allPhotos.length < 10) {
                        const varResults = await this.searchPhotos(variation, {
                            per_page: 10,
                            orientation: orientation
                        });
                        if (varResults.photos && varResults.photos.length > 0) {
                            allPhotos = allPhotos.concat(varResults.photos);
                        }
                    }
                }
            }
        }
        
        // Deduplicate and rank
        const uniquePhotos = this.deduplicatePhotos(allPhotos);
        const rankedPhotos = this.rankPhotosByRelevance(uniquePhotos, context, themeColors);
        
        if (rankedPhotos.length > 0) {
            return this.formatPhotoData(rankedPhotos[0]);
        }
        
        return null;
    }
    
    /**
     * Concept-based search using mapped terms
     */
    async conceptSearchStrategy(context, orientation) {
        const concepts = this.extractConcepts(context);
        
        if (concepts.length === 0) {
            return null;
        }
        
        console.log(`Concept search - Terms: ${concepts.join(', ')}`);
        
        for (const concept of concepts.slice(0, 3)) {
            const results = await this.searchPhotos(concept, {
                per_page: 10,
                orientation: orientation
            });
            
            if (results.photos && results.photos.length > 0) {
                const rankedPhotos = this.rankPhotosByRelevance(results.photos, context);
                if (rankedPhotos.length > 0) {
                    return this.formatPhotoData(rankedPhotos[0]);
                }
            }
        }
        
        return null;
    }
    
    /**
     * Broader search with fewer constraints
     */
    async broadSearchStrategy(context, orientation) {
        // Extract main keywords and search more broadly
        const words = context.toLowerCase().split(' ')
            .filter(w => w.length > 3)
            .slice(0, 2);
            
        const broadQuery = words.join(' ');
        
        console.log(`Broad search - Query: "${broadQuery}"`);
        
        const results = await this.searchPhotos(broadQuery, {
            per_page: 20,
            orientation: null // Remove orientation constraint
        });
        
        if (results.photos && results.photos.length > 0) {
            // Filter by orientation manually if needed
            let photos = results.photos;
            if (orientation === 'landscape') {
                photos = photos.filter(p => p.width > p.height);
            } else if (orientation === 'square') {
                photos = photos.filter(p => Math.abs(p.width - p.height) / Math.max(p.width, p.height) < 0.2);
            }
            
            if (photos.length > 0) {
                return this.formatPhotoData(photos[0]);
            }
        }
        
        return null;
    }
    
    /**
     * Curated photos fallback
     */
    async curatedFallbackStrategy(context) {
        console.log('Curated fallback - Using professional stock photos');
        
        const curatedResults = await this.getCuratedPhotos({ per_page: 30 });
        if (curatedResults && curatedResults.photos && curatedResults.photos.length > 0) {
            // Try to find somewhat relevant curated photos
            const scoredPhotos = curatedResults.photos.map(photo => ({
                ...photo,
                relevanceScore: this.calculateBasicRelevance(photo, context)
            }));
            
            const bestPhoto = scoredPhotos
                .sort((a, b) => b.relevanceScore - a.relevanceScore)
                .find(p => p.relevanceScore > 0.1);
                
            if (bestPhoto) {
                return this.formatPhotoData(bestPhoto);
            }
            
            // Return first curated photo as last resort
            return this.formatPhotoData(curatedResults.photos[0]);
        }
        
        return null;
    }
    
    /**
     * Extract concepts from context
     */
    extractConcepts(context) {
        const concepts = [];
        const contextLower = context.toLowerCase();
        
        // Use shared concept mappings
        const conceptMappings = this.getConceptMappings();
        
        // For extractConcepts, we want single best mapping per concept
        for (const [key, values] of Object.entries(conceptMappings)) {
            if (contextLower.includes(key)) {
                // Use the first (primary) mapping for each concept
                concepts.push(values[0]);
            }
        }
        
        return concepts;
    }
    
    /**
     * Remove duplicate photos by ID
     */
    deduplicatePhotos(photos) {
        const seen = new Set();
        return photos.filter(photo => {
            if (seen.has(photo.id)) {
                return false;
            }
            seen.add(photo.id);
            return true;
        });
    }
    
    /**
     * Calculate basic relevance for fallback scoring
     */
    calculateBasicRelevance(photo, context) {
        let score = 0;
        const contextWords = context.toLowerCase().split(' ');
        
        if (photo.alt) {
            const altLower = photo.alt.toLowerCase();
            contextWords.forEach(word => {
                if (word.length > 3 && altLower.includes(word)) {
                    score += 0.2;
                }
            });
        }
        
        // Prefer professional-looking images
        if (photo.alt && (photo.alt.includes('business') || photo.alt.includes('office') || photo.alt.includes('professional'))) {
            score += 0.1;
        }
        
        return Math.min(score, 1.0);
    }
    
    /**
     * Select the more specific/visual word from two words
     * @param {string} word1 - First word
     * @param {string} word2 - Second word
     * @returns {string} - More specific word
     */
    selectMoreSpecificWord(word1, word2) {
        // Common generic/abstract words
        const genericWords = new Set([
            'business', 'modern', 'professional', 'digital', 'new', 'good', 'best',
            'great', 'success', 'growth', 'future', 'global', 'quality', 'service'
        ]);
        
        // Visual/concrete words get priority
        const visualWords = new Set([
            'team', 'office', 'computer', 'laptop', 'meeting', 'desk', 'chart',
            'graph', 'people', 'person', 'building', 'city', 'nature', 'technology',
            'money', 'doctor', 'student', 'teacher', 'phone', 'tablet', 'screen'
        ]);
        
        // Check if either word is visual/concrete
        if (visualWords.has(word1.toLowerCase()) && !visualWords.has(word2.toLowerCase())) {
            return word1;
        }
        if (visualWords.has(word2.toLowerCase()) && !visualWords.has(word1.toLowerCase())) {
            return word2;
        }
        
        // Check if either word is generic
        if (genericWords.has(word1.toLowerCase()) && !genericWords.has(word2.toLowerCase())) {
            return word2;
        }
        if (genericWords.has(word2.toLowerCase()) && !genericWords.has(word1.toLowerCase())) {
            return word1;
        }
        
        // Default to the longer word (often more specific)
        return word1.length >= word2.length ? word1 : word2;
    }
    
    /**
     * Get contextual prefixes for single word queries
     * @param {string} word - Single word
     * @returns {Array<string>} - Appropriate prefixes
     */
    getContextualPrefixes(word) {
        const wordLower = word.toLowerCase();
        
        // Map words to appropriate prefixes
        if (['team', 'meeting', 'office', 'workplace'].includes(wordLower)) {
            return ['business', 'corporate', 'professional'];
        }
        if (['technology', 'computer', 'laptop', 'software'].includes(wordLower)) {
            return ['modern', 'digital', 'innovative'];
        }
        if (['growth', 'success', 'profit', 'revenue'].includes(wordLower)) {
            return ['business', 'financial', 'corporate'];
        }
        if (['education', 'learning', 'student', 'teacher'].includes(wordLower)) {
            return ['online', 'modern', 'classroom'];
        }
        if (['health', 'medical', 'patient', 'doctor'].includes(wordLower)) {
            return ['healthcare', 'medical', 'hospital'];
        }
        
        // Default prefixes for unknown words
        return ['business', 'modern', 'professional'];
    }
    
    /**
     * Clear the image cache
     */
    clearCache() {
        this.imageCache.clear();
        this.failedSearchCache.clear();
        console.log('Image cache cleared');
    }
    
    /**
     * Get detailed debugging information
     * @returns {object} - Debug information
     */
    getDebugInfo() {
        this.cleanOldRequests();
        return {
            rateLimits: this.getRateLimitStatus(),
            cacheStats: {
                successCacheSize: this.imageCache.size,
                failureCacheSize: this.failedSearchCache.size,
                totalCached: this.imageCache.size + this.failedSearchCache.size
            },
            queueStats: {
                queueLength: this.requestQueue.length,
                isProcessing: this.processing,
                requestsInHour: this.requestsInWindow.length
            },
            apiKeyConfigured: this.hasApiKey()
        };
    }
    
    /**
     * Log current system status
     */
    logStatus() {
        const debug = this.getDebugInfo();
        console.log('=== Pexels Client Status ===');
        console.log(`API Key: ${debug.apiKeyConfigured ? 'Configured' : 'Not configured'}`);
        console.log(`Rate Limits: ${debug.rateLimits.requestsInCurrentHour}/${debug.rateLimits.maxRequestsPerHour} in current hour`);
        console.log(`API Limits: ${debug.rateLimits.remaining}/${debug.rateLimits.limit} remaining`);
        console.log(`Queue: ${debug.queueStats.queueLength} pending requests`);
        console.log(`Cache: ${debug.cacheStats.successCacheSize} successful, ${debug.cacheStats.failureCacheSize} failed`);
        console.log('========================');
    }
}

// Create and export a singleton instance
const pexelsClient = new PexelsClient();