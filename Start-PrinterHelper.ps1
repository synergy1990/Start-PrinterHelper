param(
    [CmdletBinding()]
    #[Alias('Lang','L')]
    [Parameter(
        Mandatory = $false,
        Position = 1
    )]
    [ValidateSet('DE','EN')]
    [string]
    $global:Language = 'DE',

        # DNS name of the printserver
        [Parameter()]
        [string]
        $PrintServer = "SRV-PRINT"
)

switch ( $global:Language ) {
    # German translations
    'DE' {
        $CaptionDescriptionAdd               = "Einen Netzwerkdrucker hinzufügen"
        $CaptionDescriptionRemove            = "Einen lokalen Drucker entfernen"
        $CaptionDescriptionChangeDefault     = "Den Standarddrucker ändern"
        $CaptionDescriptionCancel            = "Das Programm beenden"
        $CaptionTitleWhichOption             = "Welche Option soll ausgewählt werden?"
        $CaptionChoiceAdd                    = "Netzwerkdrucker hinzufügen wurde ausgewählt."
        $CaptionChoiceRemove                 = "Lokalen Drucker entfernen wurde ausgewählt."
        $CaptionChoiceChangeDefault          = "Standarddrucker ändern wurde ausgewählt."
        $CaptionChoiceCancel                 = "Beenden wurde ausgewählt."
        $CaptionChoiceNothing                = "Es wurde nichts ausgewählt. Das Programm wird beendet."
        $CaptionTitleNetworkPrinterToBeAdded = "Welcher Netzwerkdrucker soll hinzugefügt werden?"
        $CaptionTitleLocalPrinterToBeRemoved = "Welcher lokale Drucker soll entfernt werden?"
        $CaptionTitleDefaultPrinterToSet     = "Welcher Drucker soll der neue Standarddrucker werden?"
        $CaptionChosenPrinter                = "Es wurde folgender Drucker ausgewählt: "
        $CaptionNetworkPath                  = "Der Netzwerkpfad lautet: "
    }

    # English translations
    'EN' {
        $CaptionDescriptionAdd               = "Add a network printer"
        $CaptionDescriptionRemove            = "Remove a local printer"
        $CaptionDescriptionChangeDefault     = "Change the default printer"
        $CaptionDescriptionCancel            = "Quit the program"
        $CaptionTitleWhichOption             = "Which option do you choose?"
        $CaptionChoiceAdd                    = "Choice: Add a network printer"
        $CaptionChoiceRemove                 = "Choice: Remove a local printer"
        $CaptionChoiceChangeDefault          = "Choice: Change the default printer"
        $CaptionChoiceCancel                 = "Choice: Cancel"
        $CaptionChoiceNothing                = "Choice: Nothing - program is being closed"
        $CaptionTitleNetworkPrinterToBeAdded = "Which network printer shall be added?"
        $CaptionTitleLocalPrinterToBeRemoved = "Which local printer shall be removed?"
        $CaptionTitleDefaultPrinterToSet     = "Which printer shall be the new default printer?"
        $CaptionChosenPrinter                = "Choice: "
        $CaptionNetworkPath                  = "Network path: "
    }
}

function Start-PrinterHelper {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]
        [ValidateSet( 'Add', 'Remove', 'ChangeDefault', 'Cancel' )]
        $PrinterOption = ''
    )

    $PrinterOptions = @()
    $PrinterOptions += New-Object -TypeName PSObject -Property @{
        'Description' = $CaptionDescriptionAdd
        'OptionName' = "Add"
    }

    $PrinterOptions += New-Object -TypeName PSObject -Property @{
        'Description' = $CaptionDescriptionRemove
        'OptionName' = "Remove"
    }

    $PrinterOptions += New-Object -TypeName PSObject -Property @{
        'Description' = $CaptionDescriptionChangeDefault
        'OptionName' = "ChangeDefault"
    }
    
    $PrinterOptions += New-Object -TypeName PSObject -Property @{
        'Description' = $CaptionDescriptionCancel
        'OptionName' = "Cancel"
    }
    
    while(!$PrinterOption) {
        $Choice = $PrinterOptions | Out-GridView -Title $CaptionTitleWhichOption -PassThru
    
        switch ($($Choice.OptionName)) {
            'Add'           { $PrinterOption = "Add"           ; Write-Output $CaptionChoiceAdd            ; Add-NewNetworkPrinter }
            'Remove'        { $PrinterOption = "Remove"        ; Write-Output $CaptionChoiceRemove         ; Remove-LocalPrinter }
            'ChangeDefault' { $PrinterOption = "ChangeDefault" ; Write-Output $CaptionChoiceChangeDefault  ; Set-DefaultLocalPrinter }
            'Cancel'        { $PrinterOption = "Cancel"        ; Write-Output $CaptionChoiceCancel}
            Default         { $PrinterOption = "Cancel"        ; Write-Output $CaptionChoiceNothing}
        }
    }
}

function Add-NewNetworkPrinter {
    [CmdletBinding()]
    $Printers = Get-PrintServerPrinters
    $SelectedPrinter = $Printers | Out-GridView -Title $CaptionTitleNetworkPrinterToBeAdded -PassThru
    Write-Output "$CaptionChosenPrinter $($SelectedPrinter)"
    Write-Output "$CaptionNetworkPath \\$($PrintServer)\$($SelectedPrinter.Name)"
    Add-Printer -ConnectionName "\\$($PrintServer)\$($SelectedPrinter.Name)"
}

function Remove-LocalPrinter {
    [CmdletBinding()]
    $Printers = Get-LocalPrinters
    $SelectedPrinter = $Printers | Out-GridView -Title $CaptionTitleLocalPrinterToBeRemoved -PassThru
    Write-Output "$CaptionChosenPrinter $($SelectedPrinter)"
    Write-Output "$CaptionNetworkPath \\$($PrintServer)\$($SelectedPrinter.Name)"
    Remove-Printer -Name "$($SelectedPrinter.Name)"
}

function Set-DefaultLocalPrinter {
    [CmdletBinding()]
    $Printers = Get-LocalPrinters
    $SelectedPrinter = $Printers | Out-GridView -Title $CaptionTitleDefaultPrinterToSet -PassThru
    Write-Output "$CaptionChosenPrinter $($SelectedPrinter)"
    # Select a temporary printer from "Get-CimInstance", since this is the way to mess around with the DefaultPrinterOption
    $TempPrinter = (Get-CimInstance -ClassName CIM_Printer | Where-Object {$_.Name -eq $($SelectedPrinter.Name)})
    $TempPrinter | Invoke-CimMethod -MethodName SetDefaultPrinter | Out-Null
}

function Get-PrintServerPrinters {
    [CmdletBinding()]
    # Get all the printers and return the list
    $PrintServerPrinters = Get-CimInstance Win32_Share -ComputerName $PrintServer | Where-Object { $_.Type -eq 1 }
    $PrintServerPrinters
}

function Get-LocalPrinters {
    [CmdletBinding()]
    # Get all the printers and return the list
    $LocalPrinters = Get-Printer
    $LocalPrinters
}

Start-PrinterHelper
