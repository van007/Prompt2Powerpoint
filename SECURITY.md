# Security Guidelines for Prompt2Powerpoint

## API Key Management

### Current Implementation
This application stores API keys securely in the browser's localStorage:
- **OpenRouter API Key**: Stored as `openrouter_api_key`
- **Pexels API Key**: Stored as `pexels_api_key`

### Why This is Secure
1. **No File Storage**: API keys are never written to any files in the project directory
2. **Browser-Specific**: localStorage is isolated to each user's browser
3. **No Server**: This is a client-side application with no backend to compromise

### Best Practices for Users

#### DO:
- ✅ Enter API keys only through the application's Settings interface
- ✅ Use the application only on trusted devices
- ✅ Keep your API keys private and never share them
- ✅ Regenerate API keys periodically from your provider's dashboard
- ✅ Clear browser data if using a shared computer

#### DON'T:
- ❌ Never hardcode API keys in JavaScript files
- ❌ Never create config files with real API keys
- ❌ Never commit files containing API keys
- ❌ Never share screenshots that show API keys
- ❌ Never use the same API key across multiple applications

### For Developers

#### Adding New API Integrations
If you need to add support for new APIs:

1. **Use localStorage**: Follow the existing pattern in `api.js` and `pexelsClient.js`
   ```javascript
   localStorage.setItem('your_api_key', apiKey);
   ```

2. **Never hardcode keys**: Always get keys from user input or localStorage
   ```javascript
   // Good
   const apiKey = localStorage.getItem('your_api_key');
   
   // Bad
   const apiKey = 'sk-1234567890abcdef';
   ```

3. **Update .gitignore**: Add any new config file patterns that might contain secrets

4. **Document in .env.example**: Add new environment variables to `.env.example` for reference

#### Security Checklist
Before committing code:
- [ ] No API keys in source code
- [ ] No API keys in comments
- [ ] No config files with real keys
- [ ] .gitignore includes all sensitive file patterns
- [ ] localStorage used for key storage

### Reporting Security Issues
If you discover a security vulnerability:
1. Do NOT create a public GitHub issue
2. Contact the maintainer privately
3. Allow time for a fix before disclosure

### Additional Resources
- [OpenRouter API Documentation](https://openrouter.ai/docs)
- [Pexels API Documentation](https://www.pexels.com/api/documentation/)
- [OWASP API Security](https://owasp.org/www-project-api-security/)

---

Remember: The best security practice is to assume that anything in your codebase could become public. Store secrets only in secure, external locations like browser localStorage or environment variables on a secure server.