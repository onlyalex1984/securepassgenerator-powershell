# Changelog

All notable changes to the SecurePassGenerator will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0-pre.2] - April 2025

### Added

- Check for Updates feature:
  - Automatic version detection and comparison
  - Support for both stable and pre-release versions
  - Smart installer location detection with fallback options
  - Seamless update process with minimal user interaction
  - Preservation of custom presets during updates

### Improved

- Enhanced service availability detection:
  - Improved HIBP API availability check using direct GET requests instead of ping/HEAD methods
  - Enhanced Password Pusher API detection using main website checks for more reliable connectivity
  - Better handling of corporate proxy environments
- Enhanced Installer:
  - Preservation of presets.json file during reinstallation and updates
  - Improved error handling and recovery options
  - Enhanced internet connectivity detection for corporate environments
  - Added detailed logging for troubleshooting
  - Multiple download methods with automatic fallback for improved reliability

## [1.1.0-pre.1] - April 2025

### Added

- Password Preset Management:
  - Save and load custom password presets
  - Create, edit, and delete user-defined presets
  - Enable/disable presets to customize dropdown options
  - Set default preset for application startup
  - Automatic backup of presets file when saving changes
  - Additional built-in presets (NIST Compliant, SOC 2 Compliant, Financial Compliant)

## [1.0.0] - April 2025

### Added

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
