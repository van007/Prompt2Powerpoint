# Local Images for Prompt2Powerpoint

This folder contains images that can be automatically included in your presentations when "Use Local Images" is enabled.

## How It Works

When generating presentations with local images enabled, the app will:
1. Scan all images in this folder
2. Match images to slides based on filename keywords
3. Automatically insert the best matching image for each slide

## Naming Convention

Name your images with descriptive keywords separated by hyphens (-) or underscores (_).

### Good Examples:
- `team-collaboration.jpg` - Will match slides about teamwork, collaboration
- `data-analytics-dashboard.png` - Will match slides about data, analytics, dashboards
- `innovation-technology.jpg` - Will match slides about innovation, technology
- `financial-growth-chart.png` - Will match slides about finance, growth, charts
- `customer-service.jpg` - Will match slides about customers, service

### Tips for Best Results:
1. Use descriptive keywords that match your presentation topics
2. Include multiple related keywords in the filename
3. Avoid special characters except hyphens and underscores
4. Keep filenames reasonably short but descriptive

## Supported Formats

The following image formats are supported:
- `.jpg` / `.jpeg`
- `.png`
- `.gif`
- `.webp`
- `.svg`

## Matching Algorithm

The app uses intelligent keyword matching to find the best image for each slide:
- Matches keywords from the slide title
- Matches keywords from the slide content
- Considers the overall presentation theme
- Falls back to a random image if no good match is found

## Adding Images

Simply copy your images into this folder. The app will automatically detect them the next time you generate a presentation with local images enabled.

## Notes

- If no images are found in this folder, the app will use placeholders instead
- The more descriptive your filenames, the better the matching results
- Images should be reasonably sized (recommended: under 5MB each)