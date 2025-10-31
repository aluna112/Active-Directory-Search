Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Search User"
$form.Size = New-Object System.Drawing.Size(1000,800)
$form.StartPosition = 'CenterScreen'

$global:CurrentUsers = @()

# Search by Name...
# Create the Explaination label
$explanationLabelName = New-Object System.Windows.Forms.Label
$explanationLabelName.Location = New-Object System.Drawing.Point(200, 10)
$explanationLabelName.Size = New-Object System.Drawing.Size(180, 30)
$explanationLabelName.Text = "Search by Name: "
$explanationLabelName.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabelName)
# Create Textbox
$textBoxName = New-Object System.Windows.Forms.TextBox
$textBoxName.Location = New-Object System.Drawing.Point(440,10)
$textBoxName.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxName)
# Search by EID...
# Create the Explaination label
$explanationLabelEID = New-Object System.Windows.Forms.Label
$explanationLabelEID.Location = New-Object System.Drawing.Point(650, 10)
$explanationLabelEID.Size = New-Object System.Drawing.Size(50, 30)
$explanationLabelEID.Text = "EID: "
$explanationLabelEID.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabelEID)
# Create Textbox
$textBoxEID = New-Object System.Windows.Forms.TextBox
$textBoxEID.Location = New-Object System.Drawing.Point(700,10)
$textBoxEID.Size = New-Object System.Drawing.Size(50,20)
$form.Controls.Add($textBoxEID)


# Search by Username...
# Create the Explaination label
$explanationLabelUsername = New-Object System.Windows.Forms.Label
$explanationLabelUsername.Location = New-Object System.Drawing.Point(200, 40)
$explanationLabelUsername.Size = New-Object System.Drawing.Size(200, 30)
$explanationLabelUsername.Text = "Search by Username: "
$explanationLabelUsername.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabelUsername)
# Create Textbox
$textBoxUsername = New-Object System.Windows.Forms.TextBox
$textBoxUsername.Location = New-Object System.Drawing.Point(440,40)
$textBoxUsername.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxUsername)

# Search by Title...
# Create the Explaination label
$explanationLabelTitle = New-Object System.Windows.Forms.Label
$explanationLabelTitle.Location = New-Object System.Drawing.Point(200, 70)
$explanationLabelTitle.Size = New-Object System.Drawing.Size(180, 30)
$explanationLabelTitle.Text = "Search by Title: "
$explanationLabelTitle.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabelTitle)
# Create Textbox
$textBoxTitle = New-Object System.Windows.Forms.TextBox
$textBoxTitle.Location = New-Object System.Drawing.Point(440,70)
$textBoxTitle.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxTitle)


# Search by Department...
# Create the Explaination label
$explanationLabelDept = New-Object System.Windows.Forms.Label
$explanationLabelDept.Location = New-Object System.Drawing.Point(200, 100)
$explanationLabelDept.Size = New-Object System.Drawing.Size(240, 30)
$explanationLabelDept.Text = "Search by Department: "
$explanationLabelDept.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabelDept)
# Create Textbox
$textBoxDept = New-Object System.Windows.Forms.TextBox
$textBoxDept.Location = New-Object System.Drawing.Point(440,100)
$textBoxDept.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxDept)


# Search by Service Center...
# Create the Explaination label
$explanationLabelSC = New-Object System.Windows.Forms.Label
$explanationLabelSC.Location = New-Object System.Drawing.Point(200, 130)
$explanationLabelSC.Size = New-Object System.Drawing.Size(240, 30)
$explanationLabelSC.Text = "Search by Service Center: "
$explanationLabelSC.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabelSC)
# Create Textbox
$textBoxSC = New-Object System.Windows.Forms.TextBox
$textBoxSC.Location = New-Object System.Drawing.Point(440,130)
$textBoxSC.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxSC)
# Create search button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(670,130)
$searchButton.Size = New-Object System.Drawing.Size(100,23)
$searchButton.Text = 'Search'
$searchButton.Add_Click({Get-UserInfo})
$form.Controls.Add($searchButton)

# Export button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(780, 130)
$exportButton.Size = New-Object System.Drawing.Size(100, 23)
$exportButton.Text = "Export to Excel"
$form.Controls.Add($exportButton)

# Create a single RichTextBox to display all users' information
$userInfoBox = New-Object System.Windows.Forms.RichTextBox
$userInfoBox.Location = New-Object System.Drawing.Point(20, 160)
$userInfoBox.Size = New-Object System.Drawing.Size(960, 580)
$userInfoBox.ReadOnly = $true
$userInfoBox.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($userInfoBox)

# Define a string to hold all user info
$userDetails = ""

function Get-UserInfo {
    $userDetails = "Searching..."
    $userInfoBox.Text = $userDetails

    StartSearch
}

function StartSearch {
    #$form.Controls.Remove($resultsLabel)
    #$userInfoBox.Clear()

    $global:userInputName = $textBoxName.Text
    $global:userInputUsername = $textBoxUsername.Text
    $global:userInputTitle = $textBoxTitle.Text
    $global:userInputEID = $textBoxEID.Text
    $global:userInputDept = $textBoxDept.Text
    $global:userInputSC = $textBoxSC.Text

    $noInputProvided = (
        $userInputName   -eq "" -and
        $userInputUsername -eq "" -and
        $userInputTitle    -eq "" -and
        $userInputEID   -eq "" -and
        $userInputDept -eq "" -and
        $userInputSC    -eq ""
    )
    if ($noInputProvided) {
        [System.Windows.Forms.MessageBox]::Show("No data to search. Please enter data to search...", "Search", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $userInfoBox.Clear()
        $userDetails = "Cannot blind search. Please enter data to search..."
        $userInfoBox.Text = $userDetails
        return
    }
    $filter = ""
    $filterParts = @()
    $fileNameParts = @()
    if ($userInputName -ne "") {
        $filterParts += "Name -like '*$userInputName*'"
        $fileNameParts += $userInputName
    }
    if ($userInputEID -ne "") {
        $filterParts += "EmployeeID -like '*$userInputEID*'"
        $fileNameParts += $userInputEID
    }
    if ($userInputUsername -ne "") {
        $filterParts += "SamAccountName -like '*$userInputUsername*'"
        $fileNameParts += $userInputUsername
    }
    if ($userInputTitle -ne "") {
        $filterParts += "Title -like '*$userInputTitle*'"
        $fileNameParts += $userInputTitle
    }
    if ($userInputDept -ne "") {
        $filterParts += "Department -like '*$userInputDept*'"
        $fileNameParts += $userInputDept
    }
    if ($userInputSC -ne "") {
        $filterParts += "Office -like '*$userInputSC*'"
        $fileNameParts += $userInputSC
    }
    $filter = $filterParts -join " -and "
    $global:fileName = $fileNameParts -join "+"


    write-host = $filter
    $global:CurrentUsers = Get-ADUser -Filter $filter -Properties GivenName, Surname,
        SamAccountName, UserPrincipalName, Office, City, State, Department, Title, Manager, 
        PasswordLastSet, CanonicalName, StreetAddress, telephoneNumber, EmployeeID, postalCode,
        Created, Enabled, LockedOut, Description

    $form.Controls.Remove($resultsLabel)
    $userInfoBox.Clear()

    if (!$global:CurrentUsers) {
        $userDetails = "No users found matching your search criteria."
        #return
    }

    if ($global:CurrentUsers.Count -ne $null) {
        $userDetails = "Results: ($($global:CurrentUsers.Count))`n`n"
    }

    #Loop through each user and append their info to the string
    foreach ($user in $global:CurrentUsers) {
        $userDetails += "Name: $($user.GivenName) $($user.Surname) ($($user.EmployeeID))`n"
        $userDetails += "Username: $($user.SamAccountName)`n"
        $userDetails += "Email: $($user.UserPrincipalName)`n"
        $userDetails += "Service Center: $($user.Office) ($($user.StreetAddress), $($user.City), $($user.State), $($user.postalCode))`n"
        $userDetails += "Phone/Ext: $($user.telephoneNumber)`n"
        $userDetails += "Department: $($user.Department)`n"
        $userDetails += "Title: $($user.Title)`n"

        #Clean managers name
        $managerFull = $user.Manager
        if ($managerFull) {
            $managerLong = ($managerFull -split ",")[0]
            $global:managerOnly = $managerLong.Substring(3)
            $userDetails += "Manager: $global:managerOnly"
        }

        #Get managers manager and clean managers name
        #$managersManagerFull = Get-AdUser -Filter "SamAccountName -like '$($global:managerOnly)'" -Properties Manager
        $managersManagerFull = Get-AdUser -Filter "Name -like '$($global:managerOnly)'" -Properties Manager
        if ($managersManagerFull) {
            $managersManagerLong = ($managersManagerFull.Manager -split ",")[0]
            $global:managersManagerOnly = $managersManagerLong.Substring(3)
            $userDetails += " --> $global:managersManagerOnly"
        }
        
        $userDetails += "`nLast Password Reset: $($user.PasswordLastSet)"
        $userDetails += "`nAccount Created: $($user.Created)"
        $userDetails += "`nDescription: $($user.Description)"
        $userDetails += "`nActive/Locked: $($user.Enabled)/$($user.LockedOut)"
        $userDetails += "`n"
        $userDetails += "------------------------------------------------------------"
        $userDetails += "`n"
        $userDetails += "`n"
    }

    #Display the accumulated user info in the text box
    $userInfoBox.Text += $userDetails
}

# Add option to press enter to search
$enterKeyHandler = {
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Get-UserInfo
    }
}
$textBoxes = @($textBoxName, $textBoxUsername, $textBoxTitle, $textBoxEID, $textBoxDept, $textBoxSC)
foreach ($textBox in $textBoxes) {
    $textBox.Add_KeyDown($enterKeyHandler)
}


# Export button
$exportButton.Add_Click({
    if (-not $global:CurrentUsers -or $global:CurrentUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Please perform a search first.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.FileName = "UserSearch($global:fileName).csv"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
$global:CurrentUsers | Select-Object @{
        Name='EmployeeID'; Expression={$_.EmployeeID}
    }, @{
        Name='Active'; Expression={$_.Enabled}
    }, @{
        Name='Locked Out'; Expression={$_.LockedOut}
    }, @{
        Name='First Name'; Expression={$_.GivenName}
    }, @{
        Name='Last Name'; Expression={$_.Surname}
    }, @{
        Name='Username'; Expression={$_.SamAccountName}
    }, @{
        Name='Email'; Expression={$_.UserPrincipalName}
    }, @{
        Name='Phone/EXT'; Expression={$_.telephoneNumber}
    }, @{
        Name='Service Center'; Expression={$_.Office}
    }, @{
        Name='Street Address'; Expression={$_.StreetAddress}
    }, @{
        Name='City'; Expression={$_.City}
    }, @{
        Name='State'; Expression={$_.State}
    }, @{
        Name='Department'; Expression={$_.Department}
    }, @{
        Name='Job Title'; Expression={$_.Title}
    }, @{
        Name='Manager'; Expression={
            $managerFull = $_.Manager
            if ($managerFull) {
                $managerLong = ($managerFull -split ",")[0]
                $:managerOnly = $managerLong.Substring(3)
                $managerOnly
            }
        }
    }, @{
        Name='Last Password Reset'; Expression={$_.PasswordLastSet}
    }, @{
        Name='Account Created'; Expression={$_.Created}
    }, @{
        Name='Description'; Expression={$_.Description}
    } |
    Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("Export successful!", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting file: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Show the form
$form.ShowDialog() | Out-Null