@{
    # Module manifest for PSSomeGUIThings

    # Script module associated with this manifest
    RootModule        = 'PSSomeGUIThings.psm1'

    # Version number of this module
    ModuleVersion     = '1.0.0'

    # ID used to uniquely identify this module
    GUID              = '0d4887fa-cb75-4fb0-99da-d19c8d4514e3'

    # Author of this module
    Author            = 'Loïc Ade'

    # Description of the functionality provided by this module
    Description       = 'GUI utilities for PowerShell: WPF dialogs, loading windows, window automation, theming support, and credential prompts.'

    # Minimum version of PowerShell required by this module
    PowerShellVersion = '5.1'

    # Functions to export from this module
    FunctionsToExport = '*'

    # Cmdlets to export from this module
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport  = @()

    # Aliases to export from this module
    AliasesToExport    = @()

    # Private data to pass to the module specified in RootModule
    PrivateData       = @{
        PSData = @{
            Tags       = @('GUI', 'WPF', 'UIAutomation', 'Dialog', 'Windows', 'Theme')
            ProjectUri = ''
        }
    }
}
