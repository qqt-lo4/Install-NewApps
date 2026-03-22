function New-WPFMessageBox {
    <#
    .SYNOPSIS
        Creates a customizable themed WPF message box

    .DESCRIPTION
        Wrapper around Show-WPFButtonDialog that provides standard message box button types
        with localization support. Simplifies creating common dialog scenarios.

    .PARAMETER Content
        The message text to display.

    .PARAMETER Title
        Window title.

    .PARAMETER ButtonType
        Predefined button set (OK, OK-Cancel, Yes-No, etc.).

    .PARAMETER CustomButtons
        Array of custom button hashtables with 'text' and 'value' properties.

    .PARAMETER Icon
        Icon type: None, Information, Warning, Error, Question.

    .OUTPUTS
        [String]. Value of the clicked button, or $null if closed.

    .EXAMPLE
        New-WPFMessageBox -Content "This is a message" -Title "Info" -ButtonType "OK"

    .EXAMPLE
        New-WPFMessageBox -Content "Choose an option" -ButtonType "Yes-No" -Icon Question

    .EXAMPLE
        $customButtons = @(
            @{text="Accept"; value="accept"},
            @{text="Decline"; value="decline"}
        )
        $result = New-WPFMessageBox -Content "Terms?" -ButtonType "None" -CustomButtons $customButtons

    .NOTES
        Author  : Loïc Ade
        Version : 2.0.0
        Wrapper around Show-WPFButtonDialog with localization support.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Content,

        [Parameter(Mandatory=$false, Position=1)]
        [string]$Title = "",

        [Parameter(Mandatory=$false, Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [string]$ButtonType = 'OK',

        [Parameter(Mandatory=$false, Position=3)]
        [array]$CustomButtons,

        [Parameter(Mandatory=$false, Position=4)]
        [ValidateSet('None','Information','Warning','Error','Question')]
        [string]$Icon = 'None'
    )

    # Helper function to get localized string with fallback
    function Get-ButtonText {
        param([string]$Key, [string]$Default)

        if (Get-Command Get-LocalizedString -ErrorAction SilentlyContinue) {
            try {
                $translated = Get-LocalizedString -Key $Key -ErrorAction SilentlyContinue
                if ($translated -and $translated -ne $Key) {
                    return $translated
                }
            }
            catch {
                # Fallback to default
            }
        }
        return $Default
    }

    # Define button texts with localization support
    $okText = Get-ButtonText "MessageBox.Buttons.OK" "OK"
    $cancelText = Get-ButtonText "MessageBox.Buttons.Cancel" "Cancel"
    $yesText = Get-ButtonText "MessageBox.Buttons.Yes" "Yes"
    $noText = Get-ButtonText "MessageBox.Buttons.No" "No"
    $abortText = Get-ButtonText "MessageBox.Buttons.Abort" "Abort"
    $retryText = Get-ButtonText "MessageBox.Buttons.Retry" "Retry"
    $ignoreText = Get-ButtonText "MessageBox.Buttons.Ignore" "Ignore"
    $tryAgainText = Get-ButtonText "MessageBox.Buttons.TryAgain" "Try Again"
    $continueText = Get-ButtonText "MessageBox.Buttons.Continue" "Continue"

    # Determine buttons based on ButtonType
    $buttons = @()
    if ($ButtonType -eq "None" -and $CustomButtons) {
        # Use custom buttons as-is
        $buttons = $CustomButtons
    }
    else {
        $buttons = switch ($ButtonType) {
            "OK" {
                @(@{text=$okText; value="OK"})
            }
            "OK-Cancel" {
                @(
                    @{text=$okText; value="OK"},
                    @{text=$cancelText; value="Cancel"}
                )
            }
            "Abort-Retry-Ignore" {
                @(
                    @{text=$abortText; value="Abort"},
                    @{text=$retryText; value="Retry"},
                    @{text=$ignoreText; value="Ignore"}
                )
            }
            "Yes-No-Cancel" {
                @(
                    @{text=$yesText; value="Yes"},
                    @{text=$noText; value="No"},
                    @{text=$cancelText; value="Cancel"}
                )
            }
            "Yes-No" {
                @(
                    @{text=$yesText; value="Yes"},
                    @{text=$noText; value="No"}
                )
            }
            "Retry-Cancel" {
                @(
                    @{text=$retryText; value="Retry"},
                    @{text=$cancelText; value="Cancel"}
                )
            }
            "Cancel-TryAgain-Continue" {
                @(
                    @{text=$cancelText; value="Cancel"},
                    @{text=$tryAgainText; value="TryAgain"},
                    @{text=$continueText; value="Continue"}
                )
            }
            default {
                @(@{text=$okText; value="OK"})
            }
        }
    }

    # Build parameters for Show-WPFButtonDialog
    $dialogParams = @{
        Title = $Title
        Message = $Content
        Buttons = $buttons
    }

    # Add Icon if not None
    if ($Icon -ne "None") {
        $dialogParams.Icon = $Icon
    }

    # Call Show-WPFButtonDialog
    return Show-WPFButtonDialog @dialogParams
}
