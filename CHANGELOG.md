# Changelog

All notable changes to the SecurePassGenerator will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - March 2025

### Initial Release

#### Added
- Modern WPF GUI interface for password generation and sharing
- Random password generation with customizable options:
  - Adjustable password length
  - Toggle for uppercase letters, numbers, and special characters
  - Preset configurations (Medium, Strong, Very Strong)
- Memorable password generation:
  - Support for English and Swedish word lists
  - Configurable word count
  - Options for adding uppercase, numbers, and special characters
- Password strength assessment with entropy calculation
- Integration with Have I Been Pwned API to check if passwords have been compromised
- Password Pusher integration for secure password sharing:
  - Configurable expiration settings (days and views)
  - Options for viewer deletion and retrieval steps
  - Passphrase protection support
  - QR code generation for easy mobile access
- Copy functionality for both passwords and shared URLs
- Detailed logging of all operations
- Light theme UI with modern styling
- Rate limiting for API calls to prevent abuse
