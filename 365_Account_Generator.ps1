Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#Add here to expand upon list of companies
$companies = @('Company 1', 'Company 2', 'Company 3','Company 4')
#Add locations here with same formatting as below
$locationslist = @"
Location 1
Location 2
Location 3
Location 4
Etc.

"@
$locations = $locationslist.Split([Environment]::NewLine)

#Add license names
$licenses = @('Laptop Assigned','Shared Laptop','Email Only')

$form = New-Object System.Windows.Forms.Form
$form.Text = '365 Account Generator'
$form.Size = New-Object System.Drawing.Size(350,400)
$form.StartPosition = 'CenterScreen'

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(25,20)
$label1.Size = New-Object System.Drawing.Size(130,20)
$label1.Text = 'First Name:'
$form.Controls.Add($label1)

$label1a = New-Object System.Windows.Forms.Label
$label1a.Location = New-Object System.Drawing.Point(165,20)
$label1a.Size = New-Object System.Drawing.Size(130,20)
$label1a.Text = 'Last Name:'
$form.Controls.Add($label1a)

$firstName = New-Object System.Windows.Forms.TextBox
$firstName.Location = New-Object System.Drawing.Point(25,40)
$firstName.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($firstName)

$lastName = New-Object System.Windows.Forms.TextBox
$lastName.Location = New-Object System.Drawing.Point(165,40)
$lastName.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($lastName)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(25,65)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'Company:'
$form.Controls.Add($label2)

$dropdownCompany = New-Object System.Windows.Forms.Combobox 
$dropdownCompany.Location = New-Object System.Drawing.Point(25,85)
$dropdownCompany.Size = New-Object System.Drawing.Size(275,20)
$dropdownCompany.DataSource = $companies
$form.Controls.Add($dropdownCompany)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(25,110)
$label3.Size = New-Object System.Drawing.Size(280,20)
$label3.Text = 'Location:'
$form.Controls.Add($label3)

$dropdownLocations = New-Object System.Windows.Forms.Combobox 
$dropdownLocations.Location = New-Object System.Drawing.Point(25,130)
$dropdownLocations.Size = New-Object System.Drawing.Size(275,20)
$dropdownLocations.DataSource = $locations
$form.Controls.Add($dropdownLocations)

$labelJob = New-Object System.Windows.Forms.Label
$labelJob.Location = New-Object System.Drawing.Point(25,155)
$labelJob.Size = New-Object System.Drawing.Size(280,20)
$labelJob.Text = 'Job Title:'
$form.Controls.Add($labelJob)

$jobBox = New-Object System.Windows.Forms.TextBox
$jobBox.Location = New-Object System.Drawing.Point(25,175)
$jobBox.Size = New-Object System.Drawing.Size(275,20)
$form.Controls.Add($jobBox)

$labelDept = New-Object System.Windows.Forms.Label
$labelDept.Location = New-Object System.Drawing.Point(25,200)
$labelDept.Size = New-Object System.Drawing.Size(280,20)
$labelDept.Text = 'Department:'
$form.Controls.Add($labelDept)

$deptBox = New-Object System.Windows.Forms.TextBox
$deptBox.Location = New-Object System.Drawing.Point(25,220)
$deptBox.Size = New-Object System.Drawing.Size(275,20)
$form.Controls.Add($deptBox)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(25,245)
$label4.Size = New-Object System.Drawing.Size(280,20)
$label4.Text = 'License:'
$form.Controls.Add($label4)

$dropdownLicenses = New-Object System.Windows.Forms.Combobox 
$dropdownLicenses.Location = New-Object System.Drawing.Point(25,265)
$dropdownLicenses.Size = New-Object System.Drawing.Size(275,20)
$dropdownLicenses.DataSource = $licenses
$form.Controls.Add($dropdownLicenses)

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(60,295)
$label5.Size = New-Object System.Drawing.Size(220,20)
$label5.Text = 'Default Password: Password1!'
$label5.Font = 'Consolas'
$label5.BackColor = 'Red'
$label5.TextAlign = 'MiddleCenter'
$form.Controls.Add($label5)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(120,320)
$okButton.Size = New-Object System.Drawing.Size(100,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(220,320)
$cancelButton.Size = New-Object System.Drawing.Size(100,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)


$form.Topmost = $true

$form.Add_Shown({$lastName.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $company = $dropdownCompany.Text

If($company -eq 'Company 1'){
$selection = 'youremailsuffix1.com'
}
Elseif($company -eq 'Company 2'){
$selection = '@youremailsuffix2.com'
}
Elseif($company -eq 'Company 3'){
$selection = '@youremailsuffix3.com'
}
Elseif($company -eq 'Company 4'){
$selection = '@youremailsuffix4.com'
}

##User Variables##
$fname = $firstName.Text
$sname = $lastName.Text
$company = $dropdownCompany.Text
$location = $dropdownLocations.Text
$jobtitle = $jobBox.Text
$username = ($fname[0]+$sname).ToLower()
$emailhandles = @()
$dept = $deptBox.Text
Foreach ($handle in $selection){

    if($handle -eq $selection[0]){
        $emailhandles += $username+$handle }
    else {
        $emailhandles += $username+$handle }
}
##Licensing Group Variables##
$laptopGroup = '365 GROUP ID 1'
$sharedlapGroup = '365 GROUP ID 2'
$mailonlyGroup = '365 GROUP ID 3'
if($dropdownLicenses.Text -eq 'Laptop Assigned'){
    $licenseGroup = $laptopGroup
}
elseif($dropdownLicenses.Text -eq 'Shared Laptop'){
    $licenseGroup = $sharedlapGroup
}
elseif($dropdownLicenses.Text -eq 'Email Only'){
    $licenseGroup = $mailonlyGroup
}
else{$licenseGroup = ""}

Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome
$PasswordProfile = @{
    Password = 'Password1!'
}

$newuser = New-MgUser -DisplayName "$fname $sname" -PasswordProfile $PasswordProfile -AccountEnabled -GivenName "$fname" -Surname "$sname" -MailNickname "$username" -UserPrincipalName "$username@upn_company.net" -CompanyName "$company" -Department "$dept" -OfficeLocation "$location" -JobTitle "$jobTitle"
$newuserid = $newuser.ID
New-MgGroupMember -GroupID "$licenseGroup" -DirectoryObjectID "$newuserid"
Connect-ExchangeOnline
try {
    for ($i = 1; $i -le 100; $i++ ) {
        Write-Progress -Activity "Creating Account" -Status "$i% Complete:" -PercentComplete $i
        Start-Sleep -Milliseconds 600
    }
Set-Mailbox -Identity "$username" -EmailAddresses "SMTP:$emailhandles","$username@upn_company.nett" | Out-Host
}
catch {
    for ($i = 1; $i -le 100; $i++ ) {
        Write-Host 'Assigning email aliases failed. Trying again.'
        Write-Progress -Activity "Trying Alias Again" -Status "$i% Complete:" -PercentComplete $i
        Start-Sleep -Milliseconds 300
    }
catch { Write-Host 'Mail alias assignment failed. Please assign manually.'}
    
Set-Mailbox -Identity "$username" -EmailAddresses "SMTP:$emailhandles","$username@upn_company.net" | Out-Host
}
#Disconnect-MgGraph
#Disconnect-ExchangeOnline

}
