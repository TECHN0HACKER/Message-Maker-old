$host.ui.RawUI.WindowTitle = "MESSAGE MAKER"
$Destination = "$env:LOCALAPPDATA\MESSAGE MAKER\" 
if (!([system.io.directory]::Exists($Destination))){ 
    [system.io.directory]::CreateDirectory($Destination)
}
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Set-Location $env:LOCALAPPDATA
Set-Location "MESSAGE MAKER"

# Drawing form and controls
$Form_Calculator = New-Object System.Windows.Forms.Form
    $Form_Calculator.Text = "MESSAGE MAKER"
    $Form_Calculator.Size = New-Object System.Drawing.Size(272,160)
    $Form_Calculator.FormBorderStyle = "FixedDialog"
    $Form_Calculator.TopMost = $true
    $Form_Calculator.MaximizeBox = $false
    $Form_Calculator.MinimizeBox = $true
    $Form_Calculator.ControlBox = $true
    $Form_Calculator.StartPosition = "CenterScreen"
    $Form_Calculator.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Welcome!"
    $Form_Calculator.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Next"
    $button_ClickMe.Add_Click({
$Form_Calculator.ShowInTaskbar = $false
$Form_Calculator.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_Check = New-Object System.Windows.Forms.Form
    $Form_Check.Text = "MESSAGE MAKER"
    $Form_Check.Size = New-Object System.Drawing.Size(472,742)
    $Form_Check.FormBorderStyle = "FixedDialog"
    $Form_Check.TopMost = $true
    $Form_Check.MaximizeBox = $false
    $Form_Check.MinimizeBox = $true
    $Form_Check.ControlBox = $true
    $Form_Check.StartPosition = "CenterScreen"
    $Form_Check.Font = "Segoe UI"

# adding a label to my form
$label_Check = New-Object System.Windows.Forms.Label
    $label_Check.Location = New-Object System.Drawing.Size(170,170)
    $label_Check.Size = New-Object System.Drawing.Size(200,32)
    $label_Check.TextAlign = "MiddleCenter"
    $label_Check.Text = "Select the type of message box you want"
    $Form_Check.Controls.Add($label_Check) 

    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Normal"
    $button_ClickMe.Add_Click({
$Form_Check.ShowInTaskbar = $false
$Form_Check.Opacity = 0 

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_Normal = New-Object System.Windows.Forms.Form
    $Form_Normal.Text = "MESSAGE MAKER"
    $Form_Normal.Size = New-Object System.Drawing.Size(472,742)
    $Form_Normal.FormBorderStyle = "FixedDialog"
    $Form_Normal.TopMost = $false
    $Form_Normal.MaximizeBox = $false
    $Form_Normal.MinimizeBox = $true
    $Form_Normal.ControlBox = $true
    $Form_Normal.StartPosition = "CenterScreen"
    $Form_Normal.Font = "Segoe UI"

# adding a label to my form
	$label_Normal = New-Object System.Windows.Forms.Label
    $label_Normal.Location = New-Object System.Drawing.Size(170,170)
    $label_Normal.Size = New-Object System.Drawing.Size(200,32)
    $label_Normal.TextAlign = "MiddleCenter"
    $label_Normal.Text = "Select the buttons to appear on the box"
    $Form_Normal.Controls.Add($label_Normal) 

    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "OK"
    $button_ClickMe.Add_Click({
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
        
        $Form_Normal.ShowInTaskbar = $false
        $Form_Normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $true
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_next.ShowInTaskbar = $false
$Form_next.Opacity = 0

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",0,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",0,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "x=msgbox(`"$($message)`",0,`"$($title)`")" >>message.vbs
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_Normal.Controls.Add($button_ClickMe)

     # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "OK and Cancel"
    $button_ClickMe.Add_Click({ 
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
        
        $Form_Normal.ShowInTaskbar = $false
        $Form_Normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $true
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_next.ShowInTaskbar = $false
$Form_next.Opacity = 0

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",1,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",1,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "x=msgbox(`"$($message)`",1,`"$($title)`")" >>message.vbs
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_Normal.Controls.Add($button_ClickMe)

    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Abort, Retry and Ignore"
    $button_ClickMe.Add_Click({
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
        
        $Form_Normal.ShowInTaskbar = $false
        $Form_Normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $true
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_next.ShowInTaskbar = $false
$Form_next.Opacity = 0

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",2,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",2,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "x=msgbox(`"$($message)`",2,`"$($title)`")" >>message.vbs
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_Normal.Controls.Add($button_ClickMe)

	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes, No and Cancel"
    $button_ClickMe.Add_Click({
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
        
        $Form_Normal.ShowInTaskbar = $false
        $Form_Normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $true
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_next.ShowInTaskbar = $false
$Form_next.Opacity = 0

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",3,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",3,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "x=msgbox(`"$($message)`",3,`"$($title)`")" >>message.vbs
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_Normal.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes and No"
    $button_ClickMe.Add_Click({
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
        
        $Form_Normal.ShowInTaskbar = $false
        $Form_Normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $true
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_next.ShowInTaskbar = $false
$Form_next.Opacity = 0

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",4,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",4,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "x=msgbox(`"$($message)`",4,`"$($title)`")" >>message.vbs
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_Normal.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Retry and Cancel"
    $button_ClickMe.Add_Click({
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
        
        $Form_Normal.ShowInTaskbar = $false
        $Form_Normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $true
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_next.ShowInTaskbar = $false
$Form_next.Opacity = 0

# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",5,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
    Write-Output "x=msgbox(`"$($message)`",5,`"$($title)`")" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "x=msgbox(`"$($message)`",5,`"$($title)`")" >>message.vbs
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

	# show form
	$Form_next.Add_Shown({$Form_next.Activate()})
	[void] $Form_next.ShowDialog()
    })
    $Form_Normal.Controls.Add($button_ClickMe)

	# show form
	$Form_Normal.Add_Shown({$Form_Normal.Activate()})
	[void] $Form_Normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

     # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Error"
    $button_ClickMe.Add_Click({ 
        $Form_Check.ShowInTaskbar = $false
        $Form_Check.Opacity = 0 
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_Normal = New-Object System.Windows.Forms.Form
            $Form_Normal.Text = "MESSAGE MAKER"
            $Form_Normal.Size = New-Object System.Drawing.Size(472,742)
            $Form_Normal.FormBorderStyle = "FixedDialog"
            $Form_Normal.TopMost = $false
            $Form_Normal.MaximizeBox = $false
            $Form_Normal.MinimizeBox = $true
            $Form_Normal.ControlBox = $true
            $Form_Normal.StartPosition = "CenterScreen"
            $Form_Normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Normal = New-Object System.Windows.Forms.Label
            $label_Normal.Location = New-Object System.Drawing.Size(170,170)
            $label_Normal.Size = New-Object System.Drawing.Size(200,32)
            $label_Normal.TextAlign = "MiddleCenter"
            $label_Normal.Text = "Select the buttons to appear on the box"
            $Form_Normal.Controls.Add($label_Normal) 
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",16,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",16,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",16,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
             # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK and Cancel"
            $button_ClickMe.Add_Click({ 
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",17,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",17,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",17,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Abort, Retry and Ignore"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",18,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",18,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",18,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes, No and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",19,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",19,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",19,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes and No"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",20,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",20,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",20,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Retry and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",21,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",21,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",21,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # show form
            $Form_Normal.Add_Shown({$Form_Normal.Activate()})
            [void] $Form_Normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Information"
    $button_ClickMe.Add_Click({
        $Form_Check.ShowInTaskbar = $false
        $Form_Check.Opacity = 0 
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_Normal = New-Object System.Windows.Forms.Form
            $Form_Normal.Text = "MESSAGE MAKER"
            $Form_Normal.Size = New-Object System.Drawing.Size(472,742)
            $Form_Normal.FormBorderStyle = "FixedDialog"
            $Form_Normal.TopMost = $false
            $Form_Normal.MaximizeBox = $false
            $Form_Normal.MinimizeBox = $true
            $Form_Normal.ControlBox = $true
            $Form_Normal.StartPosition = "CenterScreen"
            $Form_Normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Normal = New-Object System.Windows.Forms.Label
            $label_Normal.Location = New-Object System.Drawing.Size(170,170)
            $label_Normal.Size = New-Object System.Drawing.Size(200,32)
            $label_Normal.TextAlign = "MiddleCenter"
            $label_Normal.Text = "Select the buttons to appear on the box"
            $Form_Normal.Controls.Add($label_Normal) 
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",64,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",64,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",64,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
             # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK and Cancel"
            $button_ClickMe.Add_Click({ 
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",65,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",65,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",65,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Abort, Retry and Ignore"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",66,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",66,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",66,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes, No and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",67,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",67,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",67,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes and No"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",68,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",68,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",68,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Retry and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",69,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",69,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",69,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # show form
            $Form_Normal.Add_Shown({$Form_Normal.Activate()})
            [void] $Form_Normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Question"
    $button_ClickMe.Add_Click({
        $Form_Check.ShowInTaskbar = $false
        $Form_Check.Opacity = 0 
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_Normal = New-Object System.Windows.Forms.Form
            $Form_Normal.Text = "MESSAGE MAKER"
            $Form_Normal.Size = New-Object System.Drawing.Size(472,742)
            $Form_Normal.FormBorderStyle = "FixedDialog"
            $Form_Normal.TopMost = $false
            $Form_Normal.MaximizeBox = $false
            $Form_Normal.MinimizeBox = $true
            $Form_Normal.ControlBox = $true
            $Form_Normal.StartPosition = "CenterScreen"
            $Form_Normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Normal = New-Object System.Windows.Forms.Label
            $label_Normal.Location = New-Object System.Drawing.Size(170,170)
            $label_Normal.Size = New-Object System.Drawing.Size(200,32)
            $label_Normal.TextAlign = "MiddleCenter"
            $label_Normal.Text = "Select the buttons to appear on the box"
            $Form_Normal.Controls.Add($label_Normal) 
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",32,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",32,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",32,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
             # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK and Cancel"
            $button_ClickMe.Add_Click({ 
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",33,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",33,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",33,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Abort, Retry and Ignore"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",34,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",34,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",34,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes, No and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",35,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",35,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",35,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes and No"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",36,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",36,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",36,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Retry and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",37,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",37,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",37,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # show form
            $Form_Normal.Add_Shown({$Form_Normal.Activate()})
            [void] $Form_Normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Warning"
    $button_ClickMe.Add_Click({
        $Form_Check.ShowInTaskbar = $false
        $Form_Check.Opacity = 0 
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_Normal = New-Object System.Windows.Forms.Form
            $Form_Normal.Text = "MESSAGE MAKER"
            $Form_Normal.Size = New-Object System.Drawing.Size(472,742)
            $Form_Normal.FormBorderStyle = "FixedDialog"
            $Form_Normal.TopMost = $false
            $Form_Normal.MaximizeBox = $false
            $Form_Normal.MinimizeBox = $true
            $Form_Normal.ControlBox = $true
            $Form_Normal.StartPosition = "CenterScreen"
            $Form_Normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Normal = New-Object System.Windows.Forms.Label
            $label_Normal.Location = New-Object System.Drawing.Size(170,170)
            $label_Normal.Size = New-Object System.Drawing.Size(200,32)
            $label_Normal.TextAlign = "MiddleCenter"
            $label_Normal.Text = "Select the buttons to appear on the box"
            $Form_Normal.Controls.Add($label_Normal) 
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",48,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",48,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",48,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
             # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK and Cancel"
            $button_ClickMe.Add_Click({ 
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",49,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",49,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",49,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Abort, Retry and Ignore"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",50,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",50,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",50,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes, No and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",51,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",51,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",51,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes and No"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",52,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",52,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",52,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Retry and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",53,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",53,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",53,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # show form
            $Form_Normal.Add_Shown({$Form_Normal.Activate()})
            [void] $Form_Normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Icon on the title bar"
    $button_ClickMe.Add_Click({
        $Form_Check.ShowInTaskbar = $false
        $Form_Check.Opacity = 0 
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_Normal = New-Object System.Windows.Forms.Form
            $Form_Normal.Text = "MESSAGE MAKER"
            $Form_Normal.Size = New-Object System.Drawing.Size(472,742)
            $Form_Normal.FormBorderStyle = "FixedDialog"
            $Form_Normal.TopMost = $false
            $Form_Normal.MaximizeBox = $false
            $Form_Normal.MinimizeBox = $true
            $Form_Normal.ControlBox = $true
            $Form_Normal.StartPosition = "CenterScreen"
            $Form_Normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Normal = New-Object System.Windows.Forms.Label
            $label_Normal.Location = New-Object System.Drawing.Size(170,170)
            $label_Normal.Size = New-Object System.Drawing.Size(200,32)
            $label_Normal.TextAlign = "MiddleCenter"
            $label_Normal.Text = "Select the buttons to appear on the box"
            $Form_Normal.Controls.Add($label_Normal) 
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,220)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4096,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4096,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",4096,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
             # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,260)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "OK and Cancel"
            $button_ClickMe.Add_Click({ 
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4097,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4097,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",4097,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,300)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Abort, Retry and Ignore"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4098,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4098,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",4098,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,340)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes, No and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4099,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4099,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",4099,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,380)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes and No"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4100,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4100,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",4100,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
                    # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(140,420)
            $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Retry and Cancel"
            $button_ClickMe.Add_Click({
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        
                [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title of the message box", "MESSAGE MAKER")
                
                $Form_Normal.ShowInTaskbar = $false
                $Form_Normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $true
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_next.ShowInTaskbar = $false
        $Form_next.Opacity = 0
        
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4101,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=msgbox(`"$($message)`",4101,`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=msgbox(`"$($message)`",4101,`"$($title)`")" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
            # show form
            $Form_next.Add_Shown({$Form_next.Activate()})
            [void] $Form_next.ShowDialog()
            })
            $Form_Normal.Controls.Add($button_ClickMe)
        
            # show form
            $Form_Normal.Add_Shown({$Form_Normal.Activate()})
            [void] $Form_Normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,460)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Message (Only for Windows 10 pro and above)"
    $button_ClickMe.Add_Click({
$Form_check.ShowInTaskbar = $false
$Form_check.Opacity = 0
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
         
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_normal = New-Object System.Windows.Forms.Form
    $Form_normal.Text = "MESSAGE MAKER"
    $Form_normal.Size = New-Object System.Drawing.Size(272,160)
    $Form_normal.FormBorderStyle = "FixedDialog"
    $Form_normal.TopMost = $true
    $Form_normal.MaximizeBox = $false
    $Form_normal.MinimizeBox = $false
    $Form_normal.ControlBox = $true
    $Form_normal.StartPosition = "CenterScreen"
    $Form_normal.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Do you want your message to be spoken?"
    $Form_normal.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Yes"
    $button_ClickMe.Add_Click({
$Form_normal.ShowInTaskbar = $false
$Form_normal.Opacity = 0 
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Drawing form and controls
$Form_next = New-Object System.Windows.Forms.Form
    $Form_next.Text = "MESSAGE MAKER"
    $Form_next.Size = New-Object System.Drawing.Size(272,160)
    $Form_next.FormBorderStyle = "FixedDialog"
    $Form_next.TopMost = $false
    $Form_next.MaximizeBox = $false
    $Form_next.MinimizeBox = $false
    $Form_next.ControlBox = $true
    $Form_next.StartPosition = "CenterScreen"
    $Form_next.Font = "Segoe UI"

# adding a label to my form
    $label_Calculator = New-Object System.Windows.Forms.Label
    $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
    $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
    $label_Calculator.TextAlign = "MiddleCenter"
    $label_Calculator.Text = "Select the speaker"
    $Form_next.Controls.Add($label_Calculator)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Zira"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($message)`"" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    msg * $message
    Start-Process message.vbs
    [Environment]::Exit(1)
    })
    $Form_next.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "David"
    $button_ClickMe.Add_Click({
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    Set-Location "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    Write-Output "Dim spV" >message.vbs
    Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
    Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
    Write-Output spV.Rate=$speed >>message.vbs
    Write-Output "Spv.Speak `"$($message)`"" >>message.vbs
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Start-Process message.vbs
    msg * $message
    [Environment]::Exit(1)
})
$Form_next.Controls.Add($button_ClickMe)

# show form
$Form_next.Add_Shown({$Form_next.Activate()})
[void] $Form_next.ShowDialog()
})
$Form_Normal.Controls.Add($button_ClickMe)

# add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
    $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "No"
    $button_ClickMe.Add_Click({
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
    Set-Location $env:LOCALAPPDATA
    cd "MESSAGE MAKER"
    If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
    msg * $message
    [Environment]::Exit(1)
})
$Form_Normal.Controls.Add($button_ClickMe)

# show form
$Form_Normal.Add_Shown({$Form_Normal.Activate()})
[void] $Form_Normal.ShowDialog()
$Form_normal.Controls.Add($button_ClickMe)
         # show form
         $Form_normal.Add_Shown({$Form_normal.Activate()})
         [void] $Form_normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,500)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Message"
    $button_ClickMe.Add_Click({
        $Form_check.ShowInTaskbar = $false
        $Form_check.Opacity = 0
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
                 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_normal = New-Object System.Windows.Forms.Form
            $Form_normal.Text = "MESSAGE MAKER"
            $Form_normal.Size = New-Object System.Drawing.Size(272,160)
            $Form_normal.FormBorderStyle = "FixedDialog"
            $Form_normal.TopMost = $true
            $Form_normal.MaximizeBox = $false
            $Form_normal.MinimizeBox = $false
            $Form_normal.ControlBox = $true
            $Form_normal.StartPosition = "CenterScreen"
            $Form_normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_normal.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_normal.ShowInTaskbar = $false
        $Form_normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($message)`"" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Write-Output "WScript.echo `"$($message)`"" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($message)`"" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Write-Output "WScript.echo `"$($message)`"" >>message.vbs
            Start-Process message.vbs
            [Environment]::Exit(1)
        })
        $Form_next.Controls.Add($button_ClickMe)
        
        # show form
        $Form_next.Add_Shown({$Form_next.Activate()})
        [void] $Form_next.ShowDialog()
        })
        $Form_Normal.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "WScript.echo `"$($message)`"" >>message.vbs
            start-process message.vbs    
            [Environment]::Exit(1)
        })
        $Form_Normal.Controls.Add($button_ClickMe)
        
        # show form
        $Form_Normal.Add_Shown({$Form_Normal.Activate()})
        [void] $Form_Normal.ShowDialog()
        $Form_normal.Controls.Add($button_ClickMe)
                 # show form
                 $Form_normal.Add_Shown({$Form_normal.Activate()})
                 [void] $Form_normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

    	    # add a button
$button_ClickMe = New-Object System.Windows.Forms.Button
    $button_ClickMe.Location = New-Object System.Drawing.Size(140,540)
    $button_ClickMe.Size = New-Object System.Drawing.Size(240,32)
    $button_ClickMe.TextAlign = "MiddleCenter"
    $button_ClickMe.Text = "Input Box"
    $button_ClickMe.Add_Click({
        $Form_check.ShowInTaskbar = $false
        $Form_check.Opacity = 0
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $message = [Microsoft.VisualBasic.Interaction]::InputBox("Type the message here", "MESSAGE MAKER")
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = [Microsoft.VisualBasic.Interaction]::InputBox("Type the title here", "MESSAGE MAKER")
                 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_normal = New-Object System.Windows.Forms.Form
            $Form_normal.Text = "MESSAGE MAKER"
            $Form_normal.Size = New-Object System.Drawing.Size(272,160)
            $Form_normal.FormBorderStyle = "FixedDialog"
            $Form_normal.TopMost = $true
            $Form_normal.MaximizeBox = $false
            $Form_normal.MinimizeBox = $false
            $Form_normal.ControlBox = $true
            $Form_normal.StartPosition = "CenterScreen"
            $Form_normal.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Do you want your message to be spoken?"
            $Form_normal.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Yes"
            $button_ClickMe.Add_Click({
        $Form_normal.ShowInTaskbar = $false
        $Form_normal.Opacity = 0 
        # Load required assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        # Drawing form and controls
        $Form_next = New-Object System.Windows.Forms.Form
            $Form_next.Text = "MESSAGE MAKER"
            $Form_next.Size = New-Object System.Drawing.Size(272,160)
            $Form_next.FormBorderStyle = "FixedDialog"
            $Form_next.TopMost = $false
            $Form_next.MaximizeBox = $false
            $Form_next.MinimizeBox = $false
            $Form_next.ControlBox = $true
            $Form_next.StartPosition = "CenterScreen"
            $Form_next.Font = "Segoe UI"
        
        # adding a label to my form
            $label_Calculator = New-Object System.Windows.Forms.Label
            $label_Calculator.Location = New-Object System.Drawing.Size(8,8)
            $label_Calculator.Size = New-Object System.Drawing.Size(240,32)
            $label_Calculator.TextAlign = "MiddleCenter"
            $label_Calculator.Text = "Select the speaker"
            $Form_next.Controls.Add($label_Calculator)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(8,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "Zira"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(0)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=InputBox(`"$($message)`",`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
            })
            $Form_next.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "David"
            $button_ClickMe.Add_Click({
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $speed = [Microsoft.VisualBasic.Interaction]::InputBox("Type the speaking speed [numbers can be below 0, eg; -1, -2, -3 etc]", "MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            Set-Location "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "Dim spV" >message.vbs
            Write-Output {Set spV = CreateObject("SAPI.spVoice")} >>message.vbs
            Write-Output {Set spV.Voice = spV.GetVoices.Item(1)} >>message.vbs
            Write-Output spV.Rate=$speed >>message.vbs
            Write-Output "Spv.Speak `"$($title)`, ``$($message)`"" >>message.vbs
            Write-Output "x=InputBox(`"$($message)`",`"$($title)`")" >>message.vbs
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Start-Process message.vbs
            [Environment]::Exit(1)
        })
        $Form_next.Controls.Add($button_ClickMe)
        
        # show form
        $Form_next.Add_Shown({$Form_next.Activate()})
        [void] $Form_next.ShowDialog()
        })
        $Form_Normal.Controls.Add($button_ClickMe)
        
        # add a button
        $button_ClickMe = New-Object System.Windows.Forms.Button
            $button_ClickMe.Location = New-Object System.Drawing.Size(130,80)
            $button_ClickMe.Size = New-Object System.Drawing.Size(120,32)
            $button_ClickMe.TextAlign = "MiddleCenter"
            $button_ClickMe.Text = "No"
            $button_ClickMe.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        [System.Windows.MessageBox]::Show("Done! Your message box is ready now!","MESSAGE MAKER")
            Set-Location $env:LOCALAPPDATA
            cd "MESSAGE MAKER"
            If (Test-Path -Path message.vbs) {remove-item -path message.vbs -recurse}
            Write-Output "x=InputBox(`"$($message)`",`"$($title)`")" >>message.vbs
            start-process message.vbs    
            [Environment]::Exit(1)
        })
        $Form_Normal.Controls.Add($button_ClickMe)
        
        # show form
        $Form_Normal.Add_Shown({$Form_Normal.Activate()})
        [void] $Form_Normal.ShowDialog()
        $Form_normal.Controls.Add($button_ClickMe)
                 # show form
                 $Form_normal.Add_Shown({$Form_normal.Activate()})
                 [void] $Form_normal.ShowDialog()
    })
    $Form_Check.Controls.Add($button_ClickMe)

# show form
$Form_Check.Add_Shown({$Form_Check.Activate()})
[void] $Form_Check.ShowDialog()
    })
    $Form_Calculator.Controls.Add($button_ClickMe)

# show form
$Form_Calculator.Add_Shown({$Form_Calculator.Activate()})
[void] $Form_Calculator.ShowDialog()