![ScubaGear Logo](docs/images/SCuBA%20GitHub%20Graphic%20v6-05.png)


[![GitHub Release][github-release-img]][release]
[![PSGallery Release][psgallery-release-img]][psgallery]
[![CI Pipeline][ci-pipeline-img]][ci-pipeline]
[![Functional Tests][functional-test-img]][functional-test]
[![GitHub License][github-license-img]][license]
[![GitHub Downloads][github-downloads-img]][release]
[![PSGallery Downloads][psgallery-downloads-img]][psgallery]
[![GitHub Issues][github-issues-img]][github-issues]

ScubaGear is an assessment tool that verifies that a Microsoft 365 (M365) tenant’s configuration conforms to the policies described in the Secure Cloud Business Applications ([SCuBA](https://cisa.gov/scuba)) Secure Configuration Baseline [documents](/baselines/README.md).

> **Note**: This documentation can be read using [GitHub Pages](https://cisagov.github.io/ScubaGear).

## Target Audience

ScubaGear is for M365 administrators who want to assess their tenant environments against CISA Secure Configuration Baselines.

## What's New 🆕

**Configuration UI**: ScubaGear now includes a user-friendly graphical interface for creating and managing configuration files! 

- Launch with `Invoke-SCuBAConfigAppUI -Online`
- Intuitive setup wizard for all configuration options
- Real-time validation and YAML preview
- Microsoft Graph integration for user/group selection
- Import/export existing configurations

Perfect for users who prefer visual interfaces over command-line configuration.

## Overview

ScubaGear uses a three-step process:

- **Step One** - PowerShell code queries M365 APIs for various configuration settings.
- **Step Two** - It then calls [Open Policy Agent](https://www.openpolicyagent.org) (OPA) to compare these settings against Rego security policies written per the baseline documents.
- **Step Three** - Finally, it reports the results of the comparison as HTML, JSON, and CSV.

<img src="docs/images/scuba-process.png" />

## Key Features

### 🖥️ Multiple Interfaces
- **Configuration UI**: Graphical interface for easy setup and configuration management
- **Command Line**: PowerShell cmdlets for automation and scripting
- **Configuration Files**: YAML-based configuration for repeatable assessments

### 🔒 Comprehensive Security Coverage
- **Azure Active Directory (AAD)**: Identity and access management policies
- **Microsoft Defender**: Advanced threat protection settings  
- **Exchange Online**: Email security and compliance configurations
- **OneDrive**: File sharing and data protection policies
- **Power Platform**: Low-code application security settings
- **SharePoint**: Document collaboration and access controls
- **Microsoft Teams**: Communication and meeting security policies

### 📊 Rich Reporting
- **HTML Reports**: Interactive, user-friendly compliance reports
- **JSON Output**: Machine-readable results for automation
- **CSV Export**: Spreadsheet-compatible data for analysis

### 🎯 CISA SCuBA Alignment
- Based on official [CISA SCuBA baselines](https://cisa.gov/scuba)
- Regularly updated to match the latest security recommendations
- Detailed policy mappings and explanations

## Getting Started

### Quick Start Guide

**1. Install ScubaGear**

To install ScubaGear from [PSGallery](https://www.powershellgallery.com/packages/ScubaGear), open a PowerShell 5 terminal on a Windows computer and install the module:

```powershell
# Install ScubaGear
Install-Module -Name ScubaGear
```

**2. Install Dependencies**

```powershell
# Install the minimum required dependencies
Initialize-SCuBA 
```

**3. Verify Installation**

```powershell
# Check the version
Invoke-SCuBA -Version
```

**4. Run Your First Assessment**

```powershell
# Assess all products (basic command)
Invoke-SCuBA -ProductNames *
```

### Configuration Options

ScubaGear offers multiple ways to configure your assessments:

#### Option 1: Configuration UI (Recommended for New Users)

Use the graphical configuration interface to easily create and manage your settings:

```powershell
# Launch the Configuration UI
Invoke-SCuBAConfigAppUI -Online
```

The Configuration UI provides:
- ✅ **User-friendly interface** for all configuration options
- ✅ **Real-time validation** of yaml layout
- ✅ **YAML preview** before export configurations
- ✅ **Import/Export** existing configurations
- ✅ **Microsoft Graph integration** for user/group selection

📖 **[Learn more about the Configuration UI →](docs/scubaconfigui.md)**

#### Option 2: Configuration File

Create a YAML configuration file for repeatable, automated assessments:

```powershell
# Run with a configuration file
Invoke-SCuBA -ConfigFilePath "path/to/your/config.yaml"
```

📖 **[Learn more about Configuration Files →](docs/configuration/configuration.md)**

#### Option 3: Command Line Parameters

Use command-line parameters for quick, one-time assessments:

```powershell
# Example: Assess specific products with custom output
Invoke-SCuBA -ProductNames aad,defender -OutPath "C:\Reports"
```

📖 **[See all available parameters →](docs/configuration/parameters.md)**

> **⚠️ Important**: Successfully running ScubaGear requires certain prerequisites and configuration settings. See the [Prerequisites](#prerequisites) section below for detailed setup instructions.

## Table of Contents

### 🚀 Getting Started
- [Installation](#installation)
  - [Install from PSGallery](docs/installation/psgallery.md)
  - [Download from GitHub](docs/installation/github.md)
  - [Uninstall](docs/installation/uninstall.md)
- [Prerequisites](#prerequisites)
  - [Dependencies](docs/prerequisites/dependencies.md)
  - [Required Permissions](docs/prerequisites/permissions.md)
    - [Interactive Permissions](docs/prerequisites/interactive.md)
    - [Non-Interactive Permissions](docs/prerequisites/noninteractive.md)

### ⚙️ Configuration & Usage
- [Configuration UI](docs/scubaconfigui.md) - **Graphical interface for easy setup**
- [Configuration File](docs/configuration/configuration.md) - **YAML-based configuration**
- [Parameters Reference](docs/configuration/parameters.md) - **Command-line options**

### 🏃‍♂️ Running Assessments
- [Execution Guide](docs/execution/execution.md)
- [Understanding Reports](docs/execution/reports.md)

### 🔧 Troubleshooting & Support
- [Multiple Tenants](docs/troubleshooting/tenants.md)
- [Product-Specific Issues](docs/troubleshooting/)
  - [Defender](docs/troubleshooting/defender.md)
  - [Exchange Online](docs/troubleshooting/exchange.md)
  - [Power Platform](docs/troubleshooting/power.md)
  - [Microsoft Graph](docs/troubleshooting/graph.md)
- [Network & Proxy](docs/troubleshooting/proxy.md)

### 📚 Additional Resources

- [Assumptions](docs/misc/assumptions.md)
- [Mappings](docs/misc/mappings.md)

## Project License

Unless otherwise noted, this project is distributed under the Creative Commons Zero license. With developer approval, contributions may be submitted with an alternate compatible license. If accepted, those contributions will be listed herein with the appropriate license.

[release]: https://github.com/cisagov/ScubaGear/releases
[license]: https://github.com/cisagov/ScubaGear/blob/main/LICENSE
[psgallery]: https://www.powershellgallery.com/packages/ScubaGear
[github-cicd-workflow]: https://github.com/cisagov/ScubaGear/actions/workflows/run_pipeline.yaml
[github-issues]: https://github.com/cisagov/ScubaGear/issues
[github-license-img]: https://img.shields.io/github/license/cisagov/ScubaGear
[github-release-img]: https://img.shields.io/github/v/release/cisagov/ScubaGear?label=GitHub&logo=github
[psgallery-release-img]: https://img.shields.io/powershellgallery/v/ScubaGear?logo=powershell&label=PSGallery
[ci-pipeline]: https://github.com/cisagov/ScubaGear/actions/workflows/run_pipeline.yaml
[ci-pipeline-img]: https://github.com/cisagov/ScubaGear/actions/workflows/run_pipeline.yaml/badge.svg
[functional-test]: https://github.com/cisagov/ScubaGear/actions/workflows/test_production_function.yaml
[functional-test-img]: https://github.com/cisagov/ScubaGear/actions/workflows/test_production_function.yaml/badge.svg
[github-cicd-workflow-img]: https://img.shields.io/github/actions/workflow/status/cisagov/ScubaGear/run_pipeline.yaml?logo=github
[github-downloads-img]: https://img.shields.io/github/downloads/cisagov/ScubaGear/total?logo=github
[psgallery-downloads-img]: https://img.shields.io/powershellgallery/dt/ScubaGear?logo=powershell
[github-issues-img]: https://img.shields.io/github/issues/cisagov/ScubaGear
