# Changelog

All notable changes to the SecurePassGenerator will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - April 2025

### Initial Release

#### Added

- Modern WPF GUI interface with light theme styling
- Password Generation Features:

  - Random password generation:
    - Adjustable password length
    - Toggle for uppercase letters, numbers, and special characters
    - Preset configurations (Medium, Strong, Very Strong)
  - Memorable password generation:
    - Support for English and Swedish word lists
    - Configurable word count
    - Options for adding uppercase, numbers, and special characters
  - Copy functionality for generated passwords
  - Password strength assessment with entropy calculation

- Security Features:

  - Integration with Have I Been Pwned API to check if passwords have been compromised
  - Service availability detection with graceful handling

- Password Sharing Features:

  - Integration with Password Pusher for secure sharing
  - Configurable expiration settings (days and views)
  - Options for viewer deletion and retrieval steps
  - Passphrase protection support
  - QR code generation for easy mobile access
  - Password links history tracking and management
  - Manual password expiration capability
  - Multiple API methods (Direct and Curl) for improved reliability
  - Copy functionality for shared URLs

- Usability Features:
  - Phonetic pronunciation support:
    - NATO phonetic alphabet
    - Swedish military phonetic alphabet
  - Detailed logging of all operations
  - Rate limiting for API calls to prevent abuse
