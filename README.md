# Prompt 2 Powerpoint

<p align="center">
  <img src="assets/icon.png" alt="Prompt 2 Powerpoint Logo" width="120" height="120">
</p>

<p align="center">
  <!-- Modern MIT License -->
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge&logo=opensourceinitiative&logoColor=white" alt="License: MIT">
  <!-- Version -->
  <img src="https://img.shields.io/badge/Version-0.1.2-blue?style=for-the-badge" alt="Version">
  <!-- JavaScript -->
  <img src="https://img.shields.io/badge/JavaScript-ES6+-F7DF1E?style=for-the-badge&logo=javascript&logoColor" alt="JavaScript">
  <!-- HTML5 -->
  <img src="https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white" alt="HTML5">
  <!-- CSS3 -->
  <img src="https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white" alt="CSS3">
  <!-- PowerPoint Output -->
  <img src="https://img.shields.io/badge/Output-PowerPoint-B7472A?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white" alt="PowerPoint Output">
</p>

<p align="center">
  <!-- Contributions Welcome -->
  <img src="https://img.shields.io/badge/Contributions-Welcome-green?style=for-the-badge" alt="Contributions Welcome">
  <!-- PRs Welcome -->
  <img src="https://img.shields.io/badge/PRs-Welcome-green?style=for-the-badge&logo=github" alt="PRs Welcome">
  <!-- Maintained -->
  <img src="https://img.shields.io/badge/Maintained-Yes-green?style=for-the-badge" alt="Maintained">
</p>

<p align="center">
  <!-- jsDelivr -->
  <img src="https://img.shields.io/badge/CDN-jsDelivr-E84D3D?style=for-the-badge&logo=jsdelivr&logoColor=white" alt="jsDelivr CDN">
  <!-- cdnjs -->
  <img src="https://img.shields.io/badge/CDN-cdnjs-F38020?style=for-the-badge&logo=cloudflare&logoColor=white" alt="cdnjs">
  <!-- Google Fonts -->
  <img src="https://img.shields.io/badge/Fonts-Google_Fonts-4285F4?style=for-the-badge&logo=google-fonts&logoColor=white" alt="Google Fonts">
</p>

A powerful AI-driven web application that generates professional PowerPoint presentations using LLM APIs. Support for both privacy-focused local LLM servers and cloud-based OpenRouter API, with optional real image integration via Pexels.

## 🚀 Key Features

### Core Functionality
- **Multi-Language Support**: Generate presentations in English, French, German, Spanish, or Hindi (हिन्दी)
- **Multi-Provider AI Generation**: Choose between local LLM servers or OpenRouter cloud API
- **Privacy-First Options**: Keep your data completely private with local LLM processing
- **Smart Context Understanding**: Upload PDF/text files for context-aware presentations
- **Real Stock Images**: Optional integration with Pexels API for professional imagery
- **Theme System**: 6 professionally designed themes including a fully customizable option
- **Custom Logo Support**: Add your company/personal logo to all slides with flexible positioning
- **Interactive Editing**: Add custom slides at any position post-generation
- **Real-time Preview**: See your presentation as it's being created
- **Instant Download**: Export as standard PPTX files compatible with PowerPoint

### Image Generation Options
1. **No Images**: Clean, text-only presentations for maximum focus
2. **Image Placeholders**: Transparent placeholders that can be manually replaced in PowerPoint
3. **Real Images** (Experimental): AI-selected stock photos from Pexels based on slide content
4. **Local Images** (New): Use images from your computer folder, automatically matched to slides

### Image Layouts (when images are enabled)
- **Full Width**: Image at top, text content below
- **Side by Side**: Image on left, text on right
- **Text Focus**: Small image on right, more space for text
- **Background**: Full background image with semi-transparent text overlay

## 🔒 Privacy & Security

For maximum privacy, use:
- **Local LLM** via LM Studio or compatible server
- **Image Placeholders** instead of real images
- All processing happens in your browser - no data sent to external servers (except when using cloud options)

## 📸 Screenshots

### Configuration

<p align="center">
  <img src="assets/images/1.Settings_local_LLM.png" alt="Local LLM Settings"><br>
  <b>Local LLM Configuration</b>
</p>

<p align="center">
  <img src="assets/images/2.Settings_OpenRouter_Pexels_APIs.png" alt="OpenRouter and Pexels API Settings"><br>
  <b>OpenRouter & Pexels API Setup</b>
</p>

### Presentation Creation

<p align="center">
  <img src="assets/images/3.Presentation_Settings.png" alt="Presentation Settings"><br>
  <b>Presentation Configuration [Complexity, Slides & Language]</b>
</p>

<p align="center">
  <img src="assets/images/13.Custom_Theme.png" alt="Custom Theme Settings"><br>
  <b>Custom Theme Configuration</b>
</p>

<p align="center">
  <img src="assets/images/4.Image_Layout.png" alt="Image Layout Options"><br>
  <b>Image Source Selection with Layout</b>
</p>

<p align="center">
  <img src="assets/images/14.Custom_Logo.png" alt="Custom Logo Settings"><br>
  <b>Custom Logo Configuration</b>
</p>

<p align="center">
  <img src="assets/images/5.PDF_Upload.png" alt="PDF Upload Feature"><br>
  <b>PDF Context Upload</b>
</p>

<p align="center">
  <img src="assets/images/6.Generation.png" alt="Generation Progress"><br>
  <b>Real-time Generation Progress</b>
</p>

### Preview & Management

<p align="center">
  <img src="assets/images/7.Slide_Preview.png" alt="Slide Preview Grid"><br>
  <b>Interactive Slide Preview</b>
</p>

### Custom Slide Addition

<p align="center">
  <img src="assets/images/8.Add_new_slide.png" alt="Add New Slide Button"><br>
  <b>Add Slide Interface</b>
</p>

<p align="center">
  <img src="assets/images/9.Single_Slide_generation.png" alt="Single Slide Generation"><br>
  <b>Custom Slide Creation</b>
</p>

<p align="center">
  <img src="assets/images/10.Added_Slide_preview.png" alt="Added Slide Preview"><br>
  <b>Newly Added Slide</b>
</p>

### Generated Presentation Example

<p align="center">
  <img src="assets/images/11.Presentation_Generated.png" alt="Final Presentation"><br>
  <b>Complete Presentation</b>
</p>

<p align="center">
  <img src="assets/images/12.Silde_preview.png" alt="Detailed Slide View"><br>
  <b>Slide Details View</b>
</p>

## 🚀 Getting Started

### Option 1: Local LLM (Recommended for Privacy)
1. **Install LM Studio** or any OpenAI-compatible local LLM server
2. Start your local server (default: `http://127.0.0.1:1234`)
3. Open `index.html` in your browser
4. The app will automatically connect to your local server

### Option 2: OpenRouter API (Recommended for Large Context)
1. Get your API key from <a href="https://openrouter.ai/keys" target="_blank">OpenRouter</a> (includes free credits)
2. Open the app and go to Settings
3. Switch to "OpenRouter" provider
4. Enter your API key
5. Select from available cloud models

### Option 3: Real Images with Pexels (Optional)
1. Get a free API key from <a href="https://www.pexels.com/api/" target="_blank">Pexels</a>
2. Add the key in Settings
3. Select "Pexels Stock Images" in the image options

### Option 4: Local Images from Your Computer (New)
1. Select "Local Images" in the image source options
2. Click "Select Image Folder" and choose a folder containing images
3. Images are matched to slides by:
   - Exact title match (e.g., "Introduction.jpg" for "Introduction" slide)
   - Slide number (e.g., "slide1.jpg", "slide_1.png")
   - Keywords from slide titles

## 📝 Usage Guide

### Creating Your First Presentation

1. **Choose Your Setup**:
   - Select provider (Local LLM or OpenRouter)
   - Choose image preference (None, Placeholders, or Real Images)
   - Pick a theme from 6 options (including Custom for full control)

2. **Provide Context**:
   - Enter a detailed prompt describing your presentation
   - **Upload PDFs** for additional context (recommended for complex topics)
   - The AI will extract key information from your files

3. **Configure Generation**:
   - **Complexity**: Simple (basic), Standard (balanced), or Detailed (comprehensive)
   - **Slide Count**: 5-20 slides
   - **Theme**: Choose from 6 options including Custom
   - **Language**: Select from English, French, German, Spanish, or Hindi
   - **Logo** (Optional): Upload PNG/JPG logo, choose position and size
   - **Model**: Select based on your provider

4. **Generate & Download**:
   - Click Generate and watch real-time progress
   - Preview slides in the interactive grid
   - Download as PPTX when complete

### Custom Slide Addition (Recommended Feature!)

Perfect for adding custom flow to your presentations:

1. **Find Insert Points**: Look for ➕ icons between slides
2. **Click to Add**: Choose exact position for new content
3. **AI-Powered Generation**: New slides maintain consistent style and context
4. **Instant Preview**: See changes immediately
5. **Seamless Integration**: Download includes all custom additions

### Custom Theme Creation

Create your own personalized theme with complete control:

1. **Select Custom Theme**: Choose "Custom" from the theme dropdown
2. **Interactive Controls**: The theme preview becomes interactive
3. **Font Selection**: Choose from 12 fonts for headings and body text
4. **Color Customization**: Click any color swatch to open a color picker
5. **Real-time Preview**: See changes instantly as you customize
6. **Persistent Settings**: Your custom theme saves automatically

### Custom Logo Integration

Add professional branding to your presentations:

1. **Upload Logo**: Support for PNG/JPG images up to 5MB
2. **Position Options**: 
   - Top Right (default) - Professional placement
   - Bottom Left - Alternative positioning
3. **Size Selection**:
   - Small (5% of slide height)
   - Medium (7% of slide height)
   - Large (10% of slide height)
4. **Automatic Aspect Ratio**: Logo proportions preserved on all slides
5. **Visual Preview**: See logo placement before generation
6. **Persistent Storage**: Logo settings saved for future use

## 💡 Recommendations

### For Best Results:
- **Use OpenRouter with PDF uploads** for presentations requiring large context
- **Use Local LLM with placeholders** for maximum privacy
- **Enable real images** only for final presentations (experimental)
- **Upload relevant PDFs** to improve content accuracy
- **Use the single slide add feature** to customize presentation flow

### Provider Comparison:
| Feature | Local LLM | OpenRouter |
|---------|-----------|------------|
| Privacy | ✅ Complete | ⚠️ Cloud-based |
| Context Length | Limited | ✅ Large |
| Speed | Varies | ✅ Fast |
| Cost | ✅ Free | Free credits + paid |
| Model Selection | Limited | ✅ Many options |

## 🛠️ Technical Details

### Architecture
- **Pure Vanilla JavaScript**: No build process, no npm required
- **Modular Design**: Clean separation of concerns across modules
- **Client-Side Processing**: All operations happen in your browser
- **CDN Dependencies**: External libraries loaded from CDNs

### Core Modules
- `app.js` - Main application controller
- `api.js` - Multi-provider LLM communication
- `fileHandler.js` - PDF/text file processing
- `presentationBuilder.js` - PPTX generation with themes
- `ui.js` - User interface management
- `pexelsClient.js` - Stock image integration

### API Integrations
1. **Local LLM Server**
   - OpenAI-compatible endpoints
   - Streaming support for real-time generation

2. **OpenRouter API**
   - Access to multiple cloud models
   - Built-in rate limiting and error handling

3. **Pexels API**
   - Smart query optimization
   - Intelligent caching system
   - Rate limit: 180 requests/hour

### External Dependencies
All loaded via CDN - no installation required:
- `pptxgenjs` v3.12.0 - PowerPoint generation
- `pdf.js` v3.11.174 - PDF parsing
- `SweetAlert2` v11 - Enhanced dialogs
- `Font Awesome` v6.4.0 - Icons

## 🧪 Test Resources

Try the application with our sample PDF:
- **<a href="assets/PDF_file_uploaded.pdf" target="_blank">Sample PDF Document</a>** - Use this file to test the PDF upload and context extraction features

Presentation generated using the sample PDF as Local LLM gemma-3-27b-it
- **<a href="assets/a_beginner_s_guide_to_large_language_models_prompt_2_powerpoint.pptx" target="_blank"> Generated Presentation</a>** - Presentation Settings: Simple/5 Slides/Professional Theme/Side-by-Side Image Placeholder Layout. Added Single Slide - Local LLM vs Cloud LLM API

### How to Test:
1. Download or directly drag the sample PDF into the upload area
2. Add a prompt like "Create a presentation based on this document"
3. The AI will extract key information and generate relevant slides

## 📋 Requirements

- Modern web browser (Chrome 80+, Firefox 75+, Safari 13+, Edge 80+)
- JavaScript enabled
- Internet connection (for CDN resources and cloud APIs)
- Optional: Local LLM server for privacy mode

## 🎯 Quick Start Examples

### Example 1: Company Overview Presentation
```
Prompt: "Create a 10-slide presentation about our sustainable technology startup"
Settings: OpenRouter, Standard complexity, Theme: Modern, Real Images
Upload: company-overview.pdf
```

### Example 2: Training Material
```
Prompt: "Design a training presentation on cybersecurity best practices"
Settings: Local LLM, Detailed complexity, Theme: Professional, Placeholders
Upload: security-guidelines.pdf
```

### Example 3: Sales Pitch
```
Prompt: "Generate a sales pitch for our new AI-powered analytics platform"
Settings: OpenRouter, Simple complexity, Theme: Corporate, Real Images
Custom slides: Add specific client case studies using the + button
```

### Example 4: Custom Branded Presentation
```
Prompt: "Create a presentation about our quarterly results"
Settings: Local LLM, Standard complexity, Theme: Custom, Placeholders
Custom theme: Brand colors (#1E3A8A primary, #F59E0B accent), Roboto font
Logo: Company logo, Top Right position, Medium size
```

## 🌟 Latest Updates

### September 5, 2025
- **Multi-Language Support**: Generate presentations in 5 languages
  - English, French, German, Spanish, and Hindi (हिन्दी with Devanagari script)
  - Language preference saved in browser storage
  - All content (titles, slides, bullet points, notes) generated in selected language
  - Context-aware translations maintain presentation quality

### August 30, 2025
- **Local Images Support**: Use images from your computer folder for presentations
  - Smart image matching based on filenames and slide titles
  - Support for JPG, PNG, GIF, WebP, BMP formats
  - Automatic base64 embedding in presentations

### Previous Updates
- **Custom Logo Feature**: Upload and position company logos on all slides
- **Custom Theme Feature**: Create personalized themes with interactive controls
- **Theme System**: 6 themes including fully customizable option
- **Single Slide Addition**: Context-aware custom slide insertion
- **Multi-Provider Support**: Seamless switching between local and cloud
- **Improved Layouts**: Image layout options including background mode

## 📁 Project Structure

```
prompt-2-powerpoint/
├── index.html           # Main application
├── css/
│   └── styles.css      # Responsive styling
├── js/
│   ├── app.js          # Application controller
│   ├── api.js          # LLM provider management
│   ├── fileHandler.js  # PDF/text processing
│   ├── presentationBuilder.js # PPTX generation
│   ├── ui.js           # Interface management
│   ├── pexelsClient.js # Image API integration
│   └── localImageHandler.js # Local image folder management
└── README.md           # This file
```

## 🤝 Contributing

This is a client-side application with no backend. To contribute:
1. Fork the repository
2. Make your changes
3. Test thoroughly with both local and cloud providers
4. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

**Note**: This application processes all data locally in your browser. When using cloud services (OpenRouter, Pexels), please review their respective privacy policies.
