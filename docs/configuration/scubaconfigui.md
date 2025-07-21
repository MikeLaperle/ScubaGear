# ScubaConfig Module

The ScubaConfig module provides a graphical user interface for creating and managing ScubaGear configuration files. This module contains the PowerShell functions and resources needed to launch the configuration UI and manage ScubaGear settings.

## Overview

ScubaConfig is a PowerShell module that includes:

- **Configuration UI**: WPF-based graphical interface for creating YAML configuration files
- **Configuration Management**: Functions for loading, validating, and exporting ScubaGear configurations
- **Localization Support**: Multi-language UI text and validation messages. **Currently only supported locale is: _en-US_**
- **Connected Support**: Simplify selection screen to pull in graph users and groups
- **Debug Capabilities**: Comprehensive debugging features

## Main Function

### Invoke-SCuBAConfigAppUI

Opens the ScubaGear Configuration UI for creating and managing configuration files.

#### Syntax

```powershell
Invoke-SCuBAConfigAppUI [[-YAMLConfigFile] <String>] [[-Language] <String>] [-Online] [[-M365Environment] <String>] [-Passthru]
```

#### Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| YAMLConfigFile | String | Path to existing YAML configuration file to import | None |
| Language | String | UI language (localization) | "en-US" |
| Online | Switch | Enable Microsoft Graph connectivity | False |
| M365Environment | String | Target M365 environment (commercial, gcc, gcchigh, dod) | "commercial" |
| Passthru | Switch | Return the configuration object | False |

#### Examples

```powershell
# Basic usage - Launch the configuration UI
Invoke-SCuBAConfigAppUI

# Launch with Graph connectivity for commercial environment
Invoke-SCuBAConfigAppUI -Online -M365Environment commercial

# Import existing configuration
Invoke-SCuBAConfigAppUI -YAMLConfigFile "C:\configs\myconfig.yaml"

# Launch and connect to graph for GCC High environment
Invoke-SCuBAConfigAppUI -Online -M365Environment gcchigh
```

## Module Files

### Core Files

- **ScubaConfig.psm1**: Main module file containing all functions and UI logic
- **ScubaConfig.psd1**: Module manifest with metadata and dependencies
- **ScubaConfigAppUI.xaml**: WPF UI definition file

### Configuration Files

- **ScubaConfig_en-US.json**: English localization and configuration settings
- Additional language files can be added following the same naming pattern

### Resource Files

- UI templates, styles, and other resources as needed

## Configuration File Structure

The `ScubaConfig_en-US.json` file contains:

```json
{
  "DebugMode": "none", //supports: None, Timer, All, UI
  "Version": "1.10.0",
  "localeContext": {
    // UI text elements
  },
  "localePlaceholder": {
    // Input field placeholder text
  },
  "defaultAdvancedSettings": {
    // Default values for advanced settings
  },
  "localeInfoMessages": {
    // Success and information messages
  },
  "localeErrorMessages": {
    // Error and validation messages
  },
  "products": {
    // defines supported product for scubagear
  },
  "M365Environment": {
    //supported tenant environments for config file
  },
  "baselines": [
    "aad": {
      // defines scubagear baselines for Entra Admin Center
    },
    "defender": {
      // defines scubagear baselines for Defender Admin Center
    },
    "exo": {
      // defines scubagear baselines for M365 Exchange Admin Center
    },
    "sharepoint": {
      // defines scubagear baselines for SharePoint Admin Center
    },
    "teams": {
      // defines scubagear baselines for Teams Admin Center
    },
    "powerbi": {
      // defines scubagear baselines for Powrbi
    },
    "powerplatform": {
      // defines scubagear baselines for PowerPlatform
    }
  ],
  "exclusionTypes": {
    // defines fields and value types per exclusion baseline
  },
   "graphQueries": {
    // defines graph queries used in UI (when online)
  }
}
```

## Debug Configuration

### Debug Modes

The UI supports multiple debug modes configured in the JSON file:

- **`none`**: No debug output (default)
- **`UI`**: Debug information shown in UI debug tab
- **`Timer`**: Timer-based debug information only
- **`All`**: Complete debug output including all events

### Enabling Debug Mode

1. Edit `ScubaConfig_en-US.json` in the module directory
2. Change `"DebugMode": "none"` to desired mode
3. Restart the UI application

Example:
```json
{
  "DebugMode": "UI",
  ...
}
```

## Features

### Configuration Management
- **Organization Settings**: Tenant information, display names, descriptions
- **Product Selection**: Choose which M365 services to assess
- **Exclusions**: Configure policy exclusions for specific users, groups, or domains
- **Annotations**: Add contextual information to policies
- **Omissions**: Skip specific policies with rationale and expiration dates
- **Advanced Settings**: Output paths, authentication, and technical parameters

### User Interface
- **Tabbed Navigation**: Organized sections for different configuration areas
- **Real-time Validation**: Input validation with immediate feedback
- **Preview Generation**: Live YAML preview before export
- **Import/Export**: Load existing configurations and save new ones
- **Graph Integration**: Browse users and groups via Microsoft Graph API

### File Operations
- **YAML Import**: Load existing ScubaGear configuration files
- **YAML Export**: Save configurations in ScubaGear-compatible format
- **Clipboard Support**: Copy configurations for use elsewhere
- **Auto-naming**: Intelligent file naming based on organization settings

## Usage Workflow

1. **Launch**: Start the UI with `Invoke-SCuBAConfigAppUI`
2. **Configure**: Set organization information and select products
3. **Customize**: Add exclusions, annotations, and omissions as needed
4. **Advanced**: Configure authentication and output settings
5. **Preview**: Generate and review the YAML configuration
6. **Export**: Save or copy the configuration for use with ScubaGear

## Integration with ScubaGear

The configurations created by this UI are fully compatible with the main ScubaGear assessment tool:

```powershell
# Use the generated configuration
Invoke-SCuBA -ConfigFilePath "path\to\generated\config.yaml"
```

## Requirements

- **PowerShell 5.1** or later
- **.NET Framework 4.5** or later
- **Windows OS** with WPF support
- **ScubaGear Module** (parent module)

## Troubleshooting

### Common Issues

**UI won't launch**: Check PowerShell execution policy and .NET Framework version
**Graph connectivity fails**: Verify internet connection and authentication credentials
**Configuration validation errors**: Review required fields and format requirements

### Debug Information

Enable debug mode to get detailed information about:
- UI events and user interactions
- Configuration validation results
- Import/export operations
- Graph API calls and responses

## Development

### Extending the UI

The UI is built using WPF and follows MVVM-like patterns:
- **View**: Defined in `ScubaConfigAppUI.xaml`
- **Logic**: Contained in `ScubaConfig.psm1`
- **Data**: Managed through PowerShell hashtables and objects

### Adding Localization

1. Create new JSON file following naming pattern: `ScubaConfig_<locale>.json`
2. Translate all text elements in the localeContext section
3. Update module to load appropriate locale file

### Contributing

Follow the main ScubaGear contribution guidelines when making changes to this module.

## Version History

- **1.10.0**: Current version with full UI functionality
- Previous versions: See main ScubaGear changelog

## License

Same license as the parent ScubaGear project.

## Support

For issues and questions:
- **ScubaGear Issues**: [GitHub Issues](https://github.com/cisagov/ScubaGear/issues)
- **Documentation**: [ScubaGear Docs](https://github.com/cisagov/ScubaGear/docs)
- **Discussions**: [GitHub Discussions](https://github.com/cisagov/ScubaGear/discussions)