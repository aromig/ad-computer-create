<#
.NOTES
    ad-computer-create.ps1
    Utility for quickly creating computer accounts in Active Directory
    Companion file, config.xml, specifies default values
.SYNOPSIS
    Creates computer accounts in Active Directory
#>

#region Form Objects

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Main Form Window

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = '391,393'
$MainForm.text                   = "AD Computer Create"
$MainForm.TopMost                = $false
$MainForm.FormBorderStyle        = "Fixed3D"
$MainForm.MaximizeBox            = $false
$MainForm.ShowIcon               = $false

$rdoSingle                       = New-Object system.Windows.Forms.RadioButton
$rdoSingle.text                  = "Single Computer Account"
$rdoSingle.AutoSize              = $false
$rdoSingle.width                 = 360
$rdoSingle.height                = 20
$rdoSingle.location              = New-Object System.Drawing.Point(10,15)
$rdoSingle.Font                  = 'Microsoft Sans Serif,10'
$rdoSingle.TabIndex              = 0

$rdoMultiple                     = New-Object system.Windows.Forms.RadioButton
$rdoMultiple.text                = "Multiple Computer Accounts"
$rdoMultiple.AutoSize            = $false
$rdoMultiple.width               = 360
$rdoMultiple.height              = 20
$rdoMultiple.location            = New-Object System.Drawing.Point(10,40)
$rdoMultiple.Font                = 'Microsoft Sans Serif,10'
$rdoMultiple.TabIndex            = 1

$btnCreate                       = New-Object system.Windows.Forms.Button
$btnCreate.text                  = "Create Computer Account"
$btnCreate.width                 = 250
$btnCreate.height                = 30
$btnCreate.location              = New-Object System.Drawing.Point(70,346)
$btnCreate.Font                  = 'Microsoft Sans Serif,12'
$btnCreate.TabIndex              = 4

$progressBar                     = New-Object system.Windows.Forms.ProgressBar
$progressBar.width               = 360
$progressBar.height              = 5
$progressBar.value               = 0
$progressBar.location            = New-Object System.Drawing.Point(10,332)

#endregion Main Form Window

#region Single Machine Panel Objects

$pnlSingle                       = New-Object system.Windows.Forms.Panel
$pnlSingle.height                = 251
$pnlSingle.width                 = 360
$pnlSingle.location              = New-Object System.Drawing.Point(10,75)
$pnlSingle.TabIndex              = 2

$lblComputerName                 = New-Object system.Windows.Forms.Label
$lblComputerName.text            = "Computer Name:"
$lblComputerName.AutoSize        = $true
$lblComputerName.width           = 25
$lblComputerName.height          = 10
$lblComputerName.location        = New-Object System.Drawing.Point(10,10)
$lblComputerName.Font            = 'Microsoft Sans Serif,10'

$txtComputerName                 = New-Object system.Windows.Forms.TextBox
$txtComputerName.multiline       = $false
$txtComputerName.width           = 225
$txtComputerName.height          = 20
$txtComputerName.location        = New-Object System.Drawing.Point(120,5)
$txtComputerName.Font            = 'Microsoft Sans Serif,10'
$txtComputerName.TabIndex        = .0

$lblUserName                     = New-Object system.Windows.Forms.Label
$lblUserName.text                = "User`'s Name:"
$lblUserName.AutoSize            = $true
$lblUserName.width               = 25
$lblUserName.height              = 10
$lblUserName.location            = New-Object System.Drawing.Point(10,40)
$lblUserName.Font                = 'Microsoft Sans Serif,10'

$txtUserName                     = New-Object system.Windows.Forms.TextBox
$txtUserName.multiline           = $false
$txtUserName.width               = 225
$txtUserName.height              = 20
$txtUserName.location            = New-Object System.Drawing.Point(120,35)
$txtUserName.Font                = 'Microsoft Sans Serif,10'
$txtUserName.TabIndex            = .1

$lblSingleLocation               = New-Object system.Windows.Forms.Label
$lblSingleLocation.text          = "Location:"
$lblSingleLocation.AutoSize      = $true
$lblSingleLocation.width         = 25
$lblSingleLocation.height        = 10
$lblSingleLocation.location      = New-Object System.Drawing.Point(10,70)
$lblSingleLocation.Font          = 'Microsoft Sans Serif,10'

$txtSingleLocation               = New-Object system.Windows.Forms.TextBox
$txtSingleLocation.multiline     = $false
$txtSingleLocation.width         = 50
$txtSingleLocation.height        = 20
$txtSingleLocation.location      = New-Object System.Drawing.Point(80,65)
$txtSingleLocation.Font          = 'Microsoft Sans Serif,10'
$txtSingleLocation.TabIndex      = .2
$txtSingleLocation.MaxLength     = 3

$rdoSingleLaptop                 = New-Object system.Windows.Forms.RadioButton
$rdoSingleLaptop.text            = "Laptop"
$rdoSingleLaptop.AutoSize        = $true
$rdoSingleLaptop.width           = 104
$rdoSingleLaptop.height          = 20
$rdoSingleLaptop.location        = New-Object System.Drawing.Point(210,70)
$rdoSingleLaptop.Font            = 'Microsoft Sans Serif,10'
$rdoSingleLaptop.TabIndex        = .3
$rdoSingleLaptop.Checked         = $true

$rdoSingleDesktop                = New-Object system.Windows.Forms.RadioButton
$rdoSingleDesktop.text           = "Desktop"
$rdoSingleDesktop.AutoSize       = $true
$rdoSingleDesktop.width          = 104
$rdoSingleDesktop.height         = 20
$rdoSingleDesktop.location       = New-Object System.Drawing.Point(275,70)
$rdoSingleDesktop.Font           = 'Microsoft Sans Serif,10'
$rdoSingleDesktop.TabIndex       = .4

$rdoSingleMac                    = New-Object System.Windows.Forms.RadioButton
$rdoSingleMac.text               = "Mac"
$rdoSingleMac.AutoSize           = $true
$rdoSingleMac.width              = 104
$rdoSingleMac.height             = 20
$rdoSingleMac.location           = New-Object System.Drawing.Point(150,70)
$rdoSingleMac.Font               = 'Microsoft Sans Serif,10'
$rdoSingleMac.TabIndex           = .5

$lblSingleOU                     = New-Object system.Windows.Forms.Label
$lblSingleOU.text                = "OU:"
$lblSingleOU.AutoSize            = $true
$lblSingleOU.width               = 25
$lblSingleOU.height              = 10
$lblSingleOU.location            = New-Object System.Drawing.Point(10,100)
$lblSingleOU.Font                = 'Microsoft Sans Serif,10'

$txtSingleOU                     = New-Object system.Windows.Forms.TextBox
$txtSingleOU.multiline           = $false
$txtSingleOU.width               = 290
$txtSingleOU.height              = 20
$txtSingleOU.location            = New-Object System.Drawing.Point(50,95)
$txtSingleOU.Font                = 'Microsoft Sans Serif,10'
$txtSingleOU.TabIndex            = .6

#endregion Single Machine Panel Objects

#region Multiple Machines Panel

$pnlMultiple                     = New-Object system.Windows.Forms.Panel
$pnlMultiple.height              = 250
$pnlMultiple.width               = 360
$pnlMultiple.location            = New-Object System.Drawing.Point(10,75)
$pnlMultiple.TabIndex            = 3

$lblFile                         = New-Object system.Windows.Forms.Label
$lblFile.text                    = "Import File: [?]"
$lblFile.AutoSize                = $true
$lblFile.width                   = 30
$lblFile.height                  = 10
$lblFile.location                = New-Object System.Drawing.Point(10,7)
$lblFile.Font                    = 'Microsoft Sans Serif,10'

$txtFilePath                     = New-Object system.Windows.Forms.TextBox
$txtFilePath.multiline           = $false
$txtFilePath.width               = 150
$txtFilePath.height              = 20
$txtFilePath.location            = New-Object System.Drawing.Point(110,5)
$txtFilePath.Font                = 'Consolas,10'
$txtFilePath.TabIndex            = .0
$txtFilePath.ReadOnly            = $true

$btnSelectFile                   = New-Object system.Windows.Forms.Button
$btnSelectFile.text              = "Browse..."
$btnSelectFile.width             = 75
$btnSelectFile.height            = 20
$btnSelectFile.location          = New-Object System.Drawing.Point(270,6)
$btnSelectFile.Font              = 'Microsoft Sans Serif,10'
$btnSelectFile.TabIndex          = .1

$lblMultipleLocation             = New-Object system.Windows.Forms.Label
$lblMultipleLocation.text        = "Location:"
$lblMultipleLocation.AutoSize    = $true
$lblMultipleLocation.width       = 25
$lblMultipleLocation.height      = 10
$lblMultipleLocation.location    = New-Object System.Drawing.Point(10,36)
$lblMultipleLocation.Font        = 'Microsoft Sans Serif,10'

$txtMultipleLocation             = New-Object system.Windows.Forms.TextBox
$txtMultipleLocation.multiline   = $false
$txtMultipleLocation.width       = 50
$txtMultipleLocation.height      = 20
$txtMultipleLocation.location    = New-Object System.Drawing.Point(80,30)
$txtMultipleLocation.Font        = 'Microsoft Sans Serif,10'
$txtMultipleLocation.TabIndex    = .2
$txtMultipleLocation.MaxLength   = 3

$rdoMultipleLaptop               = New-Object system.Windows.Forms.RadioButton
$rdoMultipleLaptop.text          = "Laptop"
$rdoMultipleLaptop.AutoSize      = $true
$rdoMultipleLaptop.width         = 104
$rdoMultipleLaptop.height        = 20
$rdoMultipleLaptop.location      = New-Object System.Drawing.Point(210,35)
$rdoMultipleLaptop.Font          = 'Microsoft Sans Serif,10'
$rdoMultipleLaptop.Checked       = $true
$rdoMultipleLaptop.TabIndex      = .3

$rdoMultipleDesktop              = New-Object system.Windows.Forms.RadioButton
$rdoMultipleDesktop.text         = "Desktop"
$rdoMultipleDesktop.AutoSize     = $true
$rdoMultipleDesktop.width        = 104
$rdoMultipleDesktop.height       = 20
$rdoMultipleDesktop.location     = New-Object System.Drawing.Point(275,35)
$rdoMultipleDesktop.Font         = 'Microsoft Sans Serif,10'
$rdoMultipleDesktop.TabIndex     = .4

$lblMultipleOU                   = New-Object system.Windows.Forms.Label
$lblMultipleOU.text              = "OU:"
$lblMultipleOU.AutoSize          = $true
$lblMultipleOU.width             = 25
$lblMultipleOU.height            = 10
$lblMultipleOU.location          = New-Object System.Drawing.Point(10,65)
$lblMultipleOU.Font              = 'Microsoft Sans Serif,10'

$txtMultipleOU                   = New-Object system.Windows.Forms.TextBox
$txtMultipleOU.multiline         = $false
$txtMultipleOU.width             = 290
$txtMultipleOU.height            = 20
$txtMultipleOU.location          = New-Object System.Drawing.Point(50,60)
$txtMultipleOU.Font              = 'Microsoft Sans Serif,10'
$txtMultipleOU.TabIndex          = .5

$lblMultipleInput                = New-Object System.Windows.Forms.Label
$lblMultipleInput.text           = "Input: [?]"
$lblMultipleInput.AutoSize       = $true
$lblMultipleInput.width          = 100
$lblMultipleInput.height         = 10
$lblMultipleInput.location       = New-Object System.Drawing.Point(10, 90)
$lblMultipleInput.Font           = 'Microsoft Sans Serif,10'

$txtMultipleInput                = New-Object System.Windows.Forms.TextBox
$txtMultipleInput.multiline      = $true
$txtMultipleInput.BackColor      = "#ffffff"
$txtMultipleInput.width          = 330
$txtMultipleInput.height         = 130
$txtMultipleInput.enabled        = $true
$txtMultipleInput.location       = New-Object System.Drawing.Point(10,110)
$txtMultipleInput.TabIndex       = .6
$txtMultipleInput.Font           = 'Consolas,10'
$txtMultipleInput.ForeColor      = "#000000"
$txtMultipleInput.ScrollBars     = "Vertical"

#endregion Multiple Machines Panel

#region HelpInputs

$frmInputHelp                    = New-Object System.Windows.Forms.Form
$frmInputHelp.ClientSize         = '300,300'
$frmInputHelp.text               = 'Input Text Help'
$frmInputHelp.TopMost            = $true
$frmInputHelp.FormBorderStyle    = 'Fixed3D'
$frmInputHelp.MaximizeBox        = $false
$frmInputHelp.ShowIcon           = $false

$lblInputHelp                    = New-Object System.Windows.Forms.Label
$lblInputHelp.location           = New-Object System.Drawing.Point(10,10)
$lblInputHelp.AutoSize           = $false
$lblInputHelp.width              = 280
$lblInputHelp.height             = 280
$lblInputHelp.Font               = 'Microsoft San Serif,10'
$lblInputHelp.text += "Each line should contain a unique identifier of the computer (such as the asset tag or serial number) and the user's name, separated by a semicolon.`n`n"
$lblInputHelp.text += "Computer names will be created in the format of: [PREFIX][Location][L|D]-[ID]`n`n"
$lblInputHelp.text += "Example:`n"
$lblInputHelp.text += "  7001234;Smith,John`n"
$lblInputHelp.text += "  7005678;Doe,Jane`n`n"
$lblInputHelp.text += "could yield:`n"
$lblInputHelp.text += "  CMPDAYL-7001234`n"
$lblInputHelp.text += "  CMPDAYL-7005678`n"

$frmFileHelp                      = New-Object System.Windows.Forms.Form
$frmFileHelp.ClientSize           = '300,300'
$frmFileHelp.text                 = 'Import File Help'
$frmFileHelp.TopMost              = $true
$frmFileHelp.FormBorderStyle      = 'Fixed3D'
$frmFileHelp.MaximizeBox          = $false
$frmFileHelp.ShowIcon             = $false

$lblFileHelp                      = New-Object System.Windows.Forms.Label
$lblFileHelp.location             = New-Object System.Drawing.Point(10,10)
$lblFileHelp.AutoSize             = $false
$lblFileHelp.width                = 280
$lblFileHelp.height               = 280
$lblFileHelp.Font                 = 'Microsoft San Serif,10'
$lblFileHelp.text += "Each line should contain a unique identifier of the computer (such as the asset tag or serial number) and the user's name, separated by a semicolon.`n`n"
$lblFileHelp.text += "Computer names will be created in the format of: [PREFIX][Location][L|D]-[ID]`n`n"
$lblFileHelp.text += "Example:`n"
$lblFileHelp.text += "  7001234;Smith,John`n"
$lblFileHelp.text += "  7005678;Doe,Jane`n`n"
$lblFileHelp.text += "could yield:`n"
$lblFileHelp.text += "  CMPDAYL-7001234`n"
$lblFileHelp.text += "  CMPDAYL-7005678`n"

#endregion HelpInput

$MainForm.controls.AddRange(@($rdoSingle,$rdoMultiple,$pnlSingle,$pnlMultiple,$btnCreate,$progressBar))
$pnlSingle.controls.AddRange(@($lblComputerName,$txtComputerName,$lblUserName,$txtUserName,$lblSingleLocation,$txtSingleLocation,$rdoSingleLaptop,$rdoSingleDesktop,$rdoSingleMac,$lblSingleOU,$txtSingleOU))
$pnlMultiple.controls.AddRange(@($lblFile,$txtFilePath,$btnSelectFile,$lblMultipleLocation,$txtMultipleLocation,$rdoMultipleLaptop,$rdoMultipleDesktop,$lblMultipleOU,$txtMultipleOU,$lblMultipleInput,$txtMultipleInput))
$frmInputHelp.controls.AddRange(@($lblInputHelp))
$frmFileHelp.controls.AddRange(@($lblFileHelp))

#endregion Form Objects

#region Control Event Handlers

$rdoSingle.Add_Click({
    $pnlMultiple.Hide()
    $btnCreate.text = "Create Computer Account"
    $btnCreate.Refresh()
    $pnlSingle.Show()
    $txtComputerName.Focus()
})

$rdoMultiple.Add_Click({
    $pnlSingle.Hide()
    $btnCreate.text = "Create Computer Accounts"
    $btnCreate.Refresh()
    $pnlMultiple.Show()
    $btnSelectFile.Focus()
})

$txtSingleLocation.Add_TextChanged({
    $txtMultipleLocation.text = $txtSingleLocation.text
    $txtMultipleLocation.Refresh()
    $type = getSingleType
    $txtSingleOU.text = createOU -location $txtSingleLocation.text -type $type
})
$txtMultipleLocation.Add_TextChanged({
    $txtSingleLocation.text = $txtMultipleLocation.text
    $txtSingleLocation.Refresh()
    $type = getMultipleType
    $txtMultipleOU.text = createOU -location $txtMultipleLocation.text -type $type
})

$txtSingleOU.Add_TextChanged({
    $txtMultipleOU.text = $txtSingleOU.text
    $txtMultipleOU.Refresh()
})
$txtMultipleOU.Add_TextChanged({
    $txtSingleOU.text = $txtMultipleOU.text
    $txtSingleOU.Refresh()
})

$rdoSingleLaptop.Add_Click({
    $txtSingleOU.text = createOU -location $txtSingleLocation.text -type 'Laptops'
    $rdoMultipleLaptop.Checked = $true
})
$rdoSingleDesktop.Add_Click({
    $txtSingleOU.text = createOU -location $txtSingleLocation.text -type 'Desktops'
    $rdoMultipleDesktop.Checked = $true
})
$rdoSingleMac.Add_Click({
    $txtSingleOU.text = createOU -location $txtSingleLocation.text -type 'Macintosh'
    $rdoMultipleLaptop.Checked = $false
    $rdoMultipleDesktop.Checked = $false
})

$rdoMultipleLaptop.Add_Click({
    $txtMultipleOU.text = createOU -location $txtMultipleLocation.text -type 'Laptops'
    $rdoSingleLaptop.Checked = $true
})
$rdoMultipleDesktop.Add_Click({
    $txtMultipleOU.text = createOU -location $txtMultipleLocation.text -type 'Desktops'
    $rdoSingleDesktop.Checked = $true 
})

$btnSelectFile.Add_Click({
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('mydocuments')
        Filter = 'CSV (*.csv)|*.csv|Text (*.txt)|*.txt|All Files (*.*)|*.*'
    }
    $result = $FileBrowser.ShowDialog()
    if ($result -eq 'OK') {
        $txtFilePath.text = $FileBrowser.filename
        $txtFilePath.Refresh()

        $inputlist = Get-Content -Path $txtFilePath.text -Encoding UTF8 -Raw
        $txtMultipleInput.text = $inputlist
        $txtMultipleInput.Refresh()
    }
})

$lblMultipleInput.Add_Click({
    $frmInputHelp.ShowDialog()
})
$lblFile.Add_Click({
    $frmFileHelp.ShowDialog()
})

$btnCreate.Add_Click({
    if ($rdoSingle.Checked -eq $true) {
        if ([adsi]::Exists("LDAP://$($txtSingleOU.text)")) {
            if ($txtComputerName.text.Trim() -ne '' -and $txtSingleLocation.text.Trim() -ne '') {
                CreateSingleComputer -ComputerName $txtComputerName.text -Description $txtUserName.text -OU $txtSingleOU.text
            } else {
                if ($txtComputerName.text.Trim() -eq '') {
                    [System.Windows.Forms.MessageBox]::Show("A computer name must be specified.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                } else {
                    [System.Windows.Forms.MessageBox]::Show("A location code must be specified.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Target OU does not exist. Please correct before proceeding.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        if ([adsi]::Exists("LDAP://$($txtMultipleOU.text)")) {
            if ($txtMultipleInput.text.Trim() -eq '') {
                [System.Windows.Forms.MessageBox]::Show("An input list must be supplied. Click on [?] for details.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } elseif ($txtMultipleLocation.text.Trim() -eq '') {
                [System.Windows.Forms.MessageBox]::Show("A location code must be specified.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } elseif ($rdoMultipleLaptop.Checked -eq $false -and $rdoMultipleDesktop.Checked -eq $false) {
                [System.Windows.Forms.MessageBox]::Show("Select either Laptop or Desktop.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } else {
                CreateMultipleComputers -List $txtMultipleInput.lines -OU $txtMultipleOU.text
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Target OU does not exist. Please correct before proceeding.", "Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } 
    }
})

#endregion Control Event Handlers

#region Custom Functions

Function getSingleType {
    if ($rdoSingleLaptop.Checked -eq $true) {
        return "Laptops"
    } else {
        return "Desktops"
    }
}
function getMultipleType {
    if ($rdoMultipleLaptop.Checked -eq $true) {
        return "Laptops"
    } else {
        return "Desktops"
    }
}

Function createOU {
    param ($location, $type, $realm = $defaultRealm, $region = $defaultRegion, $container = $defaultContainer)
    $location = $location.ToUpper()

    if ($type -eq 'Macintosh') {
        $ou = "OU=$region,OU=$type,OU=$container,$realm"
    } else {
        if ($location -eq 'HBE' -or $location -eq 'AWS') {
            $location = "AWS,OU=DAY"
        }
        $ou = "OU=$location,OU=$region,OU=$type,OU=$container,$realm"
    }

    return $ou
}

Function getFQDN {
    param ($ou)

    $fqdn = ''
    $fqdn = $ou.Substring($ou.IndexOf('DC'))
    $fqdn = $fqdn.Replace('DC=','')
    $fqdn = $fqdn.Replace(',','.')

    return $fqdn
}

Function CreateSingleComputer {
    param ($computername, $description, $ou)

    UpdateProgressBar(0)

    Import-Module ActiveDirectory

    $computername = $computername.ToUpper()

    try {
        Get-ADComputer $computername -ErrorAction SilentlyContinue
        Write-Host " @ Computer Account ""$computername"" already exists in AD" -ForegroundColor Yellow
    } catch {
        try {
            New-ADComputer -Name $computername -Path $ou -Description $description -Enabled $true
            Write-Host "Created Computer Account ""$computername"" with Description ""$description""" -ForegroundColor Cyan

            if ($defaultJoinGroup -ne '' -and $rdoSingleMac.Checked -eq $false) {
                SetJoinDomainGroup -Computer $computername -Group $defaultJoinGroup
            }

            UpdateProgressBar(25)
            
            if ($rdoSingleLaptop.Checked -eq $true) {
                $computeraccount = Get-ADComputer $computername

                foreach ($laptopPolicy in $laptopPolicies) {
                    try {
                        Add-ADGroupMember -Identity $laptopPolicy -Members $computeraccount
                        Write-Host "   + Added to policy group ""$laptopPolicy""" -ForegroundColor Green
                        $progress = (($laptopPolicies.IndexOf($laptopPolicy) + 1) / $laptopPolicies.Count) * 75
                        UpdateProgressBar($progress)
                    } catch {
                        Write-Host " ! There was a problem adding $computername to policy group ""$laptopPolicy""" -ForegroundColor Red
                    }
                }

                foreach ($laptopAttribute in $laptopAttributes) {
                    if ($laptopAttribute -eq 'dnsHostName') {
                        try {
                            $dnshostname = $computername.ToLower() + '.' + $(getFQDN($txtSingleOU.text))
                            Set-ADComputer -Identity $computername -DNSHostName $dnshostname
                            Write-Host "   + Set attribute $laptopAttribute to ""$dnshostname""" -ForegroundColor Green
                        } catch {
                            Write-Host " ! There was a problem setting attribute $laptopAttribute to ""$dnshostname""" -ForegroundColor Red                       }
                    }
                }
            } elseif ($rdoSingleDesktop.Checked -eq $true) {
                $computeraccount = Get-ADComputer $computername

                foreach ($desktopPolicy in $desktopPolicies) {
                    try {
                        Add-ADGroupMember -Identity $desktopPolicy -Members $computeraccount
                        Write-Host "   + Added to policy group ""$desktopPolicy""" -ForegroundColor Green
                        $progress = (($desktopPolicies.IndexOf($desktopPolicy) + 1) / $desktopPolicies.Count) * 75
                        UpdateProgressBar($progress)
                    } catch {
                        Write-Host " ! There was a problem adding $computername to policy group ""$desktopPolicy""" -ForegroundColor Red
                    }
                }

                foreach ($desktopAttribute in $desktopAttributes) {
                    if ($desktopAttribute -eq 'dnsHostName') {
                        try {
                            $dnshostname = $computername.ToLower() + '.' + $(getFQDN($txtSingleOU.text))
                            Set-ADComputer -Identity $computername -DNSHostName $dnshostname
                            Write-Host "   + Set attribute $desktopAttribute to ""$dnshostname""" -ForegroundColor Green
                        } catch {
                            Write-Host " ! There was a problem setting attribute $desktopAttribute to ""$dnshostname""" -ForegroundColor Red
                        }
                    }
                }
            } elseif ($rdoSingleMac.Checked -eq $true) {
                $computeraccount = Get-ADComputer $computername

                foreach ($macPolicy in $macPolicies) {
                    try {
                        Add-ADGroupMember -Identity $macPolicy -Members $computeraccount
                        Write-Host "   + Added to policy group ""$macPolicy""" -ForegroundColor Green
                        $progress = (($macPolicies.IndexOf($macPolicy) + 1) / $macPolicies.Count) * 75
                        UpdateProgressBar($progress)
                    } catch {
                        Write-Host " ! There was a problem adding $computername to policy group ""$macPolicy""" -ForegroundColor Red
                    }
                }

                foreach ($macAttribute in $macAttributes) {
                    if ($macAttribute -eq 'dnsHostName') {
                        try {
                            $dnshostname = $computername.ToLower() + '.' + $(getFQDN($txtSingleOU.text))
                            Set-ADComputer -Identity $computername -DNSHostName $dnshostname
                            Write-Host "   + Set attribute $macAttribute to ""$dnshostname""" -ForegroundColor Green
                        } catch {
                            Write-Host " ! There was a problem setting attribute $macAttribute to ""$dnshostname""" -ForegroundColor Red
                        }
                    }
                }
            }

            UpdateProgressBar(100)

            [System.Windows.Forms.MessageBox]::Show("Computer Account ""$computername"" created!", "Success",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            Write-Host " ! There was a problem adding Computer Account ""$computername"" with Description ""$description""" -ForegroundColor Red
            UpdateProgressBar(0)
        }
    }
}

Function CreateMultipleComputers {
    param ($list, $ou)

    UpdateProgressBar(0)

    Import-Module ActiveDirectory

    $lines = $list.Count

    try {
        foreach ($line in $list) {
            $entry = $line -split ';'
            $id = $entry[0]
            $user = $entry[1]
            
            $machineType = ''
            if ($rdoMultipleLaptop.Checked -eq $true) {
                $machineType = 'L'
            }
            if ($rdoMultipleDesktop.Checked -eq $true) {
                $machineType = 'D'
            }

            if ($txtMultipleLocation.text -eq 'AWS') {
                $location = 'HBE'
            } else {
                $location = $txtMultipleLocation.text
            }
            
            $computername = $machinePrefix + $location + $machineType + '-' + $id
            $computername = $computername.ToUpper()

            try {
                Get-ADComputer $computername -ErrorAction SilentlyContinue
                Write-Host " @ Computer Account ""$computername"" already exists in AD" -ForegroundColor Yellow
            } catch {
                try {
                    New-ADComputer -Name $computername -Path $ou -Description $user -Enabled $true
                    Write-Host "Created Computer Account ""$computername"" with Description ""$user""" -ForegroundColor Cyan

                    if ($defaultJoinGroup -ne '') {
                        SetJoinDomainGroup -Computer $computername -Group $defaultJoinGroup
                    }
                    
                    if ($rdoMultipleLaptop.Checked -eq $true) {
                        $computeraccount = Get-ADComputer $computername
                        
                        foreach ($laptopPolicy in $laptopPolicies) {
                            try {
                                Add-ADGroupMember -Identity $laptopPolicy -Members $computeraccount
                                Write-Host "   + Added to policy group ""$laptopPolicy""" -ForegroundColor Green
                            } catch {
                                Write-Host " ! There was a problem adding $computername to policy group ""$laptopPolicy""" -ForegroundColor Red
                            }
                        }

                        foreach ($laptopAttribute in $laptopAttributes) {
                            if ($laptopAttribute -eq 'dnsHostName') {
                                try {
                                    $dnshostname = $computername.ToLower() + '.' + $(getFQDN($txtMultipleOU.text))
                                    Set-ADComputer -Identity $computername -DNSHostName $dnshostname
                                    Write-Host "   + Set attribute $laptopAttribute to ""$dnshostname""" -ForegroundColor Green
                                } catch {
                                    Write-Host " ! There was a problem setting attribute $laptopAttribute to ""$dnshostname""" -ForegroundColor Red                       }
                            }
                        }
                    } elseif ($rdoMultipleDesktop.Checked -eq $true) {
                        $computeraccount = Get-ADComputer $computername
                        
                        foreach ($desktopPolicy in $desktopPolicies) {
                            try {
                                Add-ADGroupMember -Identity $desktopPolicy -Members $computeraccount
                                Write-Host "   + Added to policy group ""$desktopPolicy""" -ForegroundColor Green
                            } catch {
                                Write-Host " ! There was a problem adding $computername to policy group ""$desktopPolicy""" -ForegroundColor Red
                            }
                        }

                        foreach ($desktopAttribute in $desktopAttributes) {
                            if ($desktopAttribute -eq 'dnsHostName') {
                                try {
                                    $dnshostname = $computername.ToLower() + '.' + $(getFQDN($txtMultipleOU.text))
                                    Set-ADComputer -Identity $computername -DNSHostName $dnshostname
                                    Write-Host "   + Set attribute $desktopAttribute to ""$dnshostname""" -ForegroundColor Green
                                } catch {
                                    Write-Host " ! There was a problem setting attribute $desktopAttribute to ""$dnshostname""" -ForegroundColor Red
                                }
                            }
                        }
                    }
                } catch {
                    Write-Host " ! There was a problem adding Computer Account ""$computername"" with Description ""$user""" -ForegroundColor Red
                }
            }
            $progress = (($list.IndexOf($line) + 1) / $lines) * 100
            UpdateProgressBar($progress)
        }

        [System.Windows.Forms.MessageBox]::Show("Computer Accounts created!", "Success",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        Write-Host " ! There was a problem iterating through the provided list." -ForegroundColor Red
        UpdateProgressBar(0)
    }
}

Function UpdateProgressBar {
    param ($percent)

    $progressBar.value = $percent
}

Function SetJoinDomainGroup {
    param ($computer, $group)

    $done = $false
    Write-Host "Waiting for Computer Account to be available" -NoNewline
    do {
        Start-Sleep -Milliseconds 500
        Write-Host "." -NoNewline

        try {
            $groupSID = (Get-ADGroup $group).sid
            $computerDN = (Get-ADComputer $computer -ErrorAction SilentlyContinue).distinguishedname
            $computerACL = Get-ACL "AD:$computerDN" -ErrorAction SilentlyContinue

            $rule1 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($groupSID, 'DeleteTree, ExtendedRight, Delete, GenericRead', 'Allow', [GUID]'00000000-0000-0000-0000-000000000000')
            $rule2 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID,'WriteProperty', 'Allow', [GUID]'4c164200-20c0-11d0-a768-00aa006e0529')
            $rule3 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID, 'Self', 'Allow', [GUID]'f3a64788-5306-11d1-a9c5-0000f80367c1')
            $rule4 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID, 'Self', 'Allow', [GUID]'72e39547-7b18-11d1-adef-00c04fd8d5cd')
            $rule5 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID,'WriteProperty', 'Allow', [GUID]'3e0abfd0-126a-11d0-a060-00aa006c33ed')
            $rule6 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID,'WriteProperty', 'Allow', [GUID]'bf967953-0de6-11d0-a285-00aa003049e2')
            $rule7 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID,'Extendedright', 'Allow', [GUID]'5f202010-79a5-11d0-9020-00c04fc2d4cf')
            $rule8 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($groupSID,'WriteProperty', 'Allow', [GUID]'bf967953-0de6-11d0-a285-00aa003049e2')

            $computerACL.AddAccessRule($rule1)
            $computerACL.AddAccessRule($rule2)
            $computerACL.AddAccessRule($rule3)
            $computerACL.AddAccessRule($rule4)
            $computerACL.AddAccessRule($rule5)
            $computerACL.AddAccessRule($rule6)
            $computerACL.AddAccessRule($rule7)
            $computerACL.AddAccessRule($rule8)

            $computerACL | Set-ACL "AD:$computerDN"

            $done = $true

            Write-Host ""
            Write-Host "   + Permissions set for $group" -ForegroundColor Green
        } catch {
            Write-Host " ! There was a problem setting permissions for $group"
            Write-Host $_
        }
    } until ($done)
}

#endregion Custom Functions


# Set script defaults
# Will be used if config.xml import is unsuccessful

$machinePrefix      = 'CMP'
$defaultLocation    = 'DAY'
$defaultRealm       = 'DC=domain,DC=forest,DC=tld'
$defaultRegion      = 'Americas'
$defaultContainer   = 'Workstations'
$defaultJoinGroup   = ''
$laptopPolicies     = @()
$laptopAttributes   = @()
$desktopPolicies    = @()
$desktopAttributes  = @()
$macPolicies        = @()
$macAttributes      = @()

$MainFormPath       = $PSScriptRoot
$configFile         = 'config.xml'
Set-Location $MainFormPath

# Set defaults from configFile

if (Test-Path $configFile) {
    try {
        $config = [XML] (Get-Content($configFile))
        $defaultLocation = $config.ADComputerCreate.LocationCode
        $txtSingleLocation.text = $defaultLocation
        $machinePrefix = $config.ADComputerCreate.MachinePrefix

        $defaultRealm = $config.ADComputerCreate.AD.Realm
        $defaultRegion = $config.ADComputerCreate.AD.Region
        $defaultContainer = $config.ADComputerCreate.AD.Container
        $defaultJoinGroup = $config.ADComputerCreate.AD.JoinDomainGroup

        $txtSingleOU.text = createOU -location $defaultLocation -type 'Laptops'

        if ($config.ADComputerCreate.Laptops.Policies.Policy.Count -ne 0) {
            foreach ($laptopPolicy in $config.ADComputerCreate.Laptops.Policies.Policy) {
                $laptopPolicies += $laptopPolicy
            }
        }
        
        if ($config.ADComputerCreate.Desktops.Policies.Policy.Count -ne 0) {
            foreach ($desktopPolicy in $config.ADComputerCreate.Desktops.Policies.Policy) {
                $desktopPolicies += $desktopPolicy
            }
        }

        if ($config.ADComputerCreate.Macs.Policies.Policy.Count -ne 0) {
            foreach ($macPolicy in $config.ADComputerCreate.Macs.Policies.Policy) {
                $macPolicies += $macPolicy
            }
        }

        if ($config.ADComputerCreate.Laptops.Attributes.Attribute.Count -ne 0) {
            foreach ($laptopAttribute in $config.ADComputerCreate.Laptops.Attributes.Attribute) {
                $laptopAttributes += $laptopAttribute
            }
        }

        if ($config.ADComputerCreate.Desktops.Attributes.Attribute.Count -ne 0) {
            foreach ($desktopAttribute in $config.ADComputerCreate.Desktops.Attributes.Attribute) {
                $desktopAttribute += $desktopAttribute
            }
        }

        if ($config.ADComputerCreate.Macs.Attributes.Attribute.Count -ne 0) {
            foreach ($macAttribute in $config.ADComputerCreate.Macs.Attributes.Attribute) {
                $macAttributes += $macAttribute
            }
        }
    } catch {
        Write-Host " ! Failed to import $configFile. Using script defaults." -ForegroundColor Red
        $txtSingleLocation.text = $defaultLocation  
        $txtSingleOU.text = createOU -location $defaultLocation -type 'Laptops'
    }
} else {
    Write-Host " ! Unable to find $configFile. Using script defaults." -ForegroundColor Red
    $txtSingleLocation.text = $defaultLocation
    $txtSingleOU.text = createOU -location $defaultLocation -type 'Laptops'
}

$pnlSingle.Hide()
$pnlMultiple.Hide()
$rdoSingle.Checked = $true

# Go!

[void]$MainForm.ShowDialog()

$MainForm.Dispose()