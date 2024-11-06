<#
.EXEInformation
.Title = "Install Printers"
.Description = ""
.ProductName = "Install Printers"
.FileVersion = "1.1.0.0"
.ProductVersion = "1.1.0.0"
.CompanyName = "
.Copyright = "© 2024 Magnus Ardström"
#>
# ps2exe.ps1 .\PrinterGUI.ps1 .\PrinterGUI.exe -noConsole -title "Install Printers" -description "Install Printers" -version "1.1.0.0" -company "na" -copyright "© 2024 Magnus Ardström" -DPIAware

# $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
# Set-AuthenticodeSignature -FilePath ".\PrinterGUI.exe" -Certificate $cert

######
###### Invoke-ps2exe .\PrinterGUI.ps1 .\PrinterGUI.exe -x64 -noConsole -version 1.0.0 -iconFile .\icon.ico
###### 
$clientId = ""
$tenantId = ""
$clientSecret = "" ##2026-10-03

# $Steps = 2
# $Step = 1
# Write-Progress -Activity "Hämtar en åtkomsttoken till Microsoft Teams" -Status "Progress:" -PercentComplete ($Step/$Steps*100) -Id 1
# $Step++

$tokenuri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body1 = @{
   client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

$tokenRequest = Invoke-WebRequest -Method Post -Uri $tokenuri -ContentType "application/x-www-form-urlencoded" -Body $body1 -UseBasicParsing
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Write-Progress -Activity "Hämtar lista på skrivare från Teams" -Status "Progress:" -PercentComplete ($Step/$Steps*100) -Id 1
# $Step++

$uri = "https://graph.microsoft.com/v1.0/sites/xxx/lists/xx/items?expand=fields(select=*)"
$resp = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing
$ListOfPrinters = $($($resp.content | ConvertFrom-Json).value.fields |Select-Object SkrivarNamn,IP,Drivrutin,Beskrivning,Arbetsplats)

# Läs statusen för "Låt Windows hantera standardskrivare"
$printerAutoSetting = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name LegacyDefaultPrinterMode


if ($ListOfPrinters){

    ### Get Current installed printers.
    $InstalledPrinters = get-printer -Name \\*
    $InstalledPrinters += get-printer -Name Printer*
    $InstalledPrinters = $($InstalledPrinters | Sort-Object)
    $DefaultPrinters = $($($InstalledPrinters.Name) + $($ListOfPrinters.SkrivarNamn))| Select-Object -Unique |Sort-Object
    $ListOfPrintersLB = $ListOfPrinters.SkrivarNamn
    $InstalledPrintersLB = $InstalledPrinters.Name

    if ($($InstalledPrintersLB).count -gt 0){
        $ListOfPrintersLB = Compare-Object $ListOfPrintersLB $InstalledPrintersLB -PassThru
    }

    # This base64 string holds the bytes that make up the orange 'G' icon (just an example for a 32x32 pixel image)
    $iconBase64      = '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'
    $iconBytes              = [Convert]::FromBase64String($iconBase64)
    # initialize a Memory stream holding the bytes
    $stream                 = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    $Form                               = New-Object system.Windows.Forms.Form
    $Form.ClientSize                    = New-Object System.Drawing.Point(587, 400)
    $Form.text                          = "FO Peterson - Installera skrivare v1.1"
    $Form.TopMost                       = $false
    $Form.MaximizeBox                   = $false
    $Form.MinimizeBox                   = $false
    $Form.FormBorderStyle               = 'Fixed3D'
    $Form.Icon                          = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
    
    # Skapar en Label för informativ text med radbrytning
    $InfoLabel                          = New-Object system.Windows.Forms.Label
    $InfoLabel.AutoSize                 = $false
    $InfoLabel.Text                     = "Välj en eller flera skrivare att installera från vänster kolumn, eller ta bort en skrivare från höger kolumn. Du kan också välja en standardskrivare från rullgardinsmenyn nedan. Avaktivera 'Låt Windows hantera standardskrivare' om du själv vill välja standardskrivare."
    $InfoLabel.width                    = 500
    $InfoLabel.height                   = 45
    $InfoLabel.location                 = New-Object System.Drawing.Point(34, 10)
    $InfoLabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
    
    # Justera placeringen av ListBox och andra komponenter
    $AddPrinterLB                       = New-Object system.Windows.Forms.ListBox
    $AddPrinterLB.text                  = "listBox"
    $AddPrinterLB.width                 = 252
    $AddPrinterLB.height                = 200
    $AddPrinterLB.location              = New-Object System.Drawing.Point(34, 80)
    $AddPrinterLB.SelectionMode         = 'MultiExtended'
    $AddPrinterLB.Sorted                = $true
    
    $RemPrinterLB                       = New-Object system.Windows.Forms.ListBox
    $RemPrinterLB.text                  = "listBox"
    $RemPrinterLB.width                 = 255
    $RemPrinterLB.height                = 200
    $RemPrinterLB.location              = New-Object System.Drawing.Point(300, 80)
    $RemPrinterLB.SelectionMode         = 'MultiExtended'
    $RemPrinterLB.Sorted                = $true
    
    $DefPrinterCB                       = New-Object system.Windows.Forms.ComboBox
    $DefPrinterCB.text                  = "Välj Standard skrivare:"
    $DefPrinterCB.width                 = 300
    $DefPrinterCB.height                = 13
    $DefPrinterCB.location              = New-Object System.Drawing.Point(160, 300)
    $DefPrinterCB.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    $DefPrinterCB.Sorted                = $true
    
    # Knappar
    $cancelButton                       = New-Object system.Windows.Forms.Button
    $cancelButton.text                  = "Avbryt"
    $cancelButton.width                 = 60
    $cancelButton.height                = 30
    $cancelButton.location              = New-Object System.Drawing.Point(301, 350)
    $cancelButton.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $cancelButton.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton                  = $cancelButton
    
    $okButton                           = New-Object system.Windows.Forms.Button
    $okButton.text                      = "Fortsätt"
    $okButton.width                     = 60
    $okButton.height                    = 30
    $okButton.location                  = New-Object System.Drawing.Point(226, 350)
    $okButton.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $okButton.Enabled                   = $false  # Börja med knappen inaktiverad
    $okButton.DialogResult              = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton                  = $okButton
    
    # Labels
    $LabellistBox1                      = New-Object system.Windows.Forms.Label
    $LabellistBox1.text                 = "Installera ny skrivare:"
    $LabellistBox1.AutoSize             = $true
    $LabellistBox1.width                = 25
    $LabellistBox1.height               = 10
    $LabellistBox1.location             = New-Object System.Drawing.Point(34, 60)
    $LabellistBox1.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    
    $LabellistBox2                      = New-Object system.Windows.Forms.Label
    $LabellistBox2.text                 = "Ta bort skrivare:"
    $LabellistBox2.AutoSize             = $true
    $LabellistBox2.width                = 25
    $LabellistBox2.height               = 10
    $LabellistBox2.location             = New-Object System.Drawing.Point(301, 60)
    $LabellistBox2.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    
    # Skapar en switchbox för att hantera "Låt Windows hantera standardskrivare"
    $AutoManageCheckbox = New-Object system.Windows.Forms.CheckBox
    $AutoManageCheckbox.Text = "Låt Windows hantera standardskrivare"
    $AutoManageCheckbox.Location = New-Object System.Drawing.Point(34, 330)
    $AutoManageCheckbox.Width = 300
    $AutoManageCheckbox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

# Ställ in checkboxens status baserat på registret
if ($printerAutoSetting -eq 0) {
    $AutoManageCheckbox.Checked = $true  # Windows hanterar standardskrivaren
} else {
    $AutoManageCheckbox.Checked = $false  # Användaren hanterar standardskrivaren manuellt
}

# Event handler för när användaren ändrar inställningen
$AutoManageCheckbox.add_CheckedChanged({
    if ($AutoManageCheckbox.Checked -eq $true) {
        # Skriv till registret för att låta Windows hantera standardskrivaren
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name LegacyDefaultPrinterMode -Value 0
    } else {
        # Skriv till registret för att låta användaren hantera standardskrivaren
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name LegacyDefaultPrinterMode -Value 1
    }
})

    # Händelsehanterare för att aktivera "Fortsätt"-knappen om något valts i ComboBox
    $DefPrinterCB.add_SelectedIndexChanged({
        if ($DefPrinterCB.SelectedItem -ne $null) {
            $okButton.Enabled = $true  # Aktivera knappen om ett val har gjorts
        }
        else {
            $okButton.Enabled = $false # Inaktivera om inget val finns
        }
    })
    
    # Event handler for adding selected printer from AddPrinterLB to DefPrinterCB
    $AddPrinterLB.add_SelectedIndexChanged({
        $selectedPrinter = $AddPrinterLB.SelectedItem
        if ($selectedPrinter -ne $null) {
            # Kontrollera om skrivaren redan finns i DefPrinterCB
            if (-not $DefPrinterCB.Items.Contains($selectedPrinter)) {
                [void]$DefPrinterCB.Items.Add($selectedPrinter)
            }
            # Välj automatiskt skrivaren i DefPrinterCB
            $DefPrinterCB.SelectedItem = $selectedPrinter
        }
    })


    # Lägger till alla kontroller
    $Form.controls.AddRange(@($InfoLabel, $AddPrinterLB, $RemPrinterLB, $cancelButton, $okButton, $LabellistBox1, $LabellistBox2, $DefPrinterCB,$AutoManageCheckbox))
    
    # Visa formuläret
    #[void]$Form.ShowDialog()
    
    if (($InstalledPrinters.count -gt 0)-or $ListOfPrintersLB.count -gt 0) {
        # Menu if more than one printer. 
        $i = 0
        foreach ($printer in $ListOfPrintersLB) {
            $i++
            [void] $AddPrinterLB.Items.Add($($printer))
            Remove-Variable printer
        }
        foreach ($printer in $InstalledPrintersLB) {
            $i++
            [void] $RemPrinterLB.Items.Add($($printer))
            Remove-Variable printer
        }
        foreach ($printer in $DefaultPrinters) {
            $i++
            [void] $DefPrinterCB.Items.Add($printer)
            Remove-Variable printer
        }

        $form.Controls.Add($AddPrinterLB)
        $form.Controls.Add($RemPrinterLB)
        $form.Controls.Add($DefPrinterCB)
        $form.Topmost = $true
            
        $result = $form.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # write-host "add: $($AddPrinterLB.SelectedItems)"
            # write-host "remove: $($RemPrinterLB.SelectedItems)"
            # write-host "Default: $($DefPrinterCB.SelectedItem)"

            ### Remove printers
            foreach ($printer in $($RemPrinterLB.SelectedItems)) {
                Remove-Printer -name $printer
            }

            ### Add printers
            foreach ($printer in $($AddPrinterLB.SelectedItems)) {
                $toinstall = $ListOfPrinters[$($ListOfPrinters.SkrivarNamn).IndexOf($printer)]
                $checkPortExists = Get-Printerport -Name $("IP_" + $($toinstall.IP))-ErrorAction SilentlyContinue
                if (-not $checkPortExists) {
                    Add-PrinterPort -name $("IP_" + $($toinstall.IP)) -PrinterHostAddress $($toinstall.IP)
                }
                Add-Printer -Name $($toinstall.SkrivarNamn) -DriverName $($toinstall.Drivrutin) -PortName $("IP_" + $($toinstall.IP))

            }
            
            if ($($DefPrinterCB.SelectedItem) -gt 0){
                $Default = $($DefPrinterCB.SelectedItem) 
                $DefaultPrinter = (Get-CimInstance -ClassName CIM_Printer | Where-Object {$_.Name -eq $Default}[0])
                $DefaultPrinter | Invoke-CimMethod -MethodName SetDefaultPrinter | Out-Null
            }
        }
        $stream.Dispose()
        $Form.Dispose()

        Add-Type -AssemblyName System.Windows.Forms

        # Skapa en popup med Ja och Nej-alternativ
        $result = [System.Windows.Forms.MessageBox]::Show("Installation av skrivare klar ! Vill du öppna inställningar för skrivare?", "FO Peterson - Installera skrivare v1.1", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Om användaren väljer "Yes", öppna ms-settings:printers
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process "ms-settings:printers"
        }


    }
}
