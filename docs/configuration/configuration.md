# ScubaGear Configuration File

ScubaGear uses a single consolidated configuration file to specify all parameters and settings for the `Invoke-SCuBA` cmdlet. The path to this file is provided via the `-ConfigFilePath` parameter, and the file must be formatted as YAML. All configuration for ScubaGear—including product-specific settings, policy exclusions, and metadata—should be included in this one file.

- The configuration file allows you to set or override any supported ScubaGear parameter, as well as embed additional fields for supplemental metadata.
- All variable names use Pascal case and are consistent with the parameters used by ScubaGear.
- If a parameter is specified both on the command line and in the configuration file, the command line value takes precedence.

## Sample Configuration File

A comprehensive example is provided in [`full_config.yaml`](../../PowerShell/ScubaGear/Sample-Config-Files/full_config.yaml). This file demonstrates all supported parameters and configuration options, including product-specific settings, policy exclusions, and metadata fields.

To use ScubaGear with your configuration file:

```powershell
Invoke-SCuBA -ConfigFilePath full_config.yaml
```

You may override any parameter at the command line as needed:

```powershell
Invoke-SCuBA -ConfigFilePath full_config.yaml -ProductNames defender
```

## Generating a Sample Configuration File

You can generate a template configuration file using the `New-SCuBAConfig` cmdlet. This will create a `SampleConfig.yaml` file with all available options, which you can then edit to suit your environment.

```powershell
New-SCuBAConfig
```

## Omitting Policies

To omit specific policies from ScubaGear evaluation (for example, if a policy is not applicable to your organization), use the `OmitPolicy` top-level key in your configuration file. Omitted policies will appear as "Omitted" in the HTML report and will be colored gray. For each omitted policy, you can specify:

- `Rationale`: The reason for omitting the policy (displayed in the report).
- `Expiration`: (Optional) A date after which the policy should no longer be omitted (format: yyyy-mm-dd).

**Omitting policies should only be done with appropriate risk management approval, as it may introduce blind spots in your assessment.**

<<<<<<< HEAD
## Annotate Policies

ScubaGear supports annotating results for individual policies. Annotated policies will be shown in the HTML with the
annotation appended to the details column. Annotated policies are intended to:
- Document action plans for any failed controls. ScubaGear will output a warning for any failing controls that are not
documented in the config file, though this warning can be silenced with the `-SilenceBODWarnings` flag.
- Allow users to identify incorrect results
- Help contextualize results

The `AnnotatePolicy` top-level key, shown in this [example ScubaGear configuration file](../../PowerShell/ScubaGear/Sample-Config-Files/annotate_policies.yaml), allows the user to specify the policies that should be annotated.

For each annotated policy, the config file allows you to indicate the following:
- `IncorrectResult`: Boolean, whether or not to mark the result incorrect. Optional, defaults to false.
- `Comment`: The annotation to add to the report. A warning will be printed if control is marked incorrect with no comment provided as justification.
- `RemediationDate`: Optional. The date a failing control is anticipated to be implemented. The expected format is yyy-mm-dd.

**Exercise care when marking incorrect results because this can inadvertently introduce blind spots when assessing your system.**

## Product-specific Configuration
=======
## Product-Specific Configuration
>>>>>>> 260bdea (removed old configs and updated readme)

All product-specific configuration is now included in the single configuration file. For example, Entra ID, Defender, and Exchange Online settings are specified under their respective top-level keys. See the `full_config.yaml` for examples.

### Entra ID Example

```yaml
Aad:
  MS.AAD.1.1v1:
    CapExclusions:
      Users:
        - <user-guid>
      Groups:
        - <group-guid>
```

### Defender Example

```yaml
Defender:
  MS.DEFENDER.1.4v1:
    SensitiveAccounts:
      IncludedUsers:
        - <user-guid>
      ExcludedGroups:
        - <group-guid>
```

### Exchange Online Example

```yaml
Exo:
  MS.EXO.1.1v2:
    AllowedForwardingDomains:
      - example.com
      - partner.org
```

## YAML Anchors and Aliases

YAML [anchors and aliases](https://smcleod.net/2022/11/yaml-anchors-and-aliases/) can be used within your configuration file to avoid repeating values. For example:

```yaml
CommonSensitiveAccountFilter: &CommonSensitiveAccountFilter
  IncludedUsers:
    - <user-guid>
  ExcludedGroups:
    - <group-guid>

Defender:
  MS.DEFENDER.1.4v1:
    SensitiveAccounts: *CommonSensitiveAccountFilter
  MS.DEFENDER.1.5v1:
    SensitiveAccounts: *CommonSensitiveAccountFilter
```

Using anchors and aliases is optional, but can help keep your configuration DRY and maintainable.

## Muting Version Check Warnings

To prevent ScubaGear from checking for newer releases and emitting a warning at import time, set the environment variable `SCUBAGEAR_SKIP_VERSION_CHECK` to any non-whitespace value.

---

**Note:** All configuration is now managed in a single file. Please refer to [`full_config.yaml`](../../PowerShell/ScubaGear/Sample-Config-Files/full_config.yaml) for the latest structure and options.