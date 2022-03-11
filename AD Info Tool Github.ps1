Add-Type -assembly System.Windows.Forms | Out-Null
$UnlockADuserMessage = New-Object -ComObject WScript.Shell 
$RemoveADGroupConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$ADDADGroupConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$ResetADUserPwConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$ClearAllGroupFromUserConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$DisableUserAccountConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$TerminateUserAccountConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$TerminateComputerAccountConfirm = New-Object -ComObject Wscript.Shell -ErrorAction Stop
#-----Define main window property
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = 'AD Info Tools '
$main_form.StartPosition = "CenterScreen"
$main_form.Width = 600

$main_form.Height = 400
$main_form.AutoSize = $true

$ADuserBox = New-Object System.Windows.Forms.ComboBox

$ADuserBox.Width = 300


#----Create a label element on the form

$aduser = New-Object System.Windows.Forms.Label

$aduser.Text = "AD users: "

$aduser.Location = New-Object System.Drawing.Point(5, 10)

$aduser.AutoSize = $true

$main_form.Controls.Add($aduser)
#----Create a label element on the form

$ADGroupList = New-Object System.Windows.Forms.Label

$ADGroupList.Text = "AD Group: "

$ADGroupList.Location = New-Object System.Drawing.Point(240, 290)

$ADGroupList.AutoSize = $true

$main_form.Controls.Add($ADGroupList)


#----Creat Active User Drop Down List

$ADuserBox = New-Object System.Windows.Forms.ComboBox

$ADuserBox.Width = 150

$ADuserBox.Location = New-Object System.Drawing.Point(70, 10)

$main_form.Controls.Add($ADuserBox)


$ADGroupListBox = New-Object System.Windows.Forms.ComboBox

$ADGroupListBox.Width = 200

$ADGroupListBox.Location = New-Object System.Drawing.Point(300, 290)

$main_form.Controls.Add($ADGroupListBox)

#----Creat Lockout User Drop Down List

$LockoutUsersBox = New-Object System.Windows.Forms.ComboBox

$LockoutUsersBox.Width = 150
$LockoutUsersBox.SelectedValue = "  "

$LockoutUsersBox.Location = New-Object System.Drawing.Point(120, 70)

$main_form.Controls.Add($LockoutUsersBox)

#---------Getting AD Group List
$ADGroup = Get-ADGroup -Filter * | Select-Object Name | Sort-Object Name
foreach ($Group in $ADGroup) {
    [void]$ADGroupListBox.Items.Add($Group.Name);
}



#---------Getting Active AD user info

$Users = (Get-ADUser -Filter * -Property Enabled | Where-Object { $_.Enabled -like "True" } | Select-Object Name -ExpandProperty Name | Sort-Object Name)
Foreach ($User in $Users) {

    [void]$ADuserBox.Items.Add($User.Name);

}
#---------Getting Lockout user info

$LockoutUsersList = Search-ADAccount -LockedOut

Foreach ($User in $LockoutUsersList) {

    [void]$LockoutUsersBox.Items.Add($User.SamAccountName);

}
#----------------------------------------------------------------------
#----------------------------------------------------------------------
#----------------------------------------------------------------------
#Label and Text Section

$LockoutUsers = New-Object System.Windows.Forms.Label

$LockoutUsers.Text = "Lockout Account: "

$LockoutUsers.Location = New-Object System.Drawing.Point(5, 70)

$LockoutUsers.AutoSize = $true

$main_form.Controls.Add($LockoutUsers)

#----------------------------------------------------------------------
#show the time of the last password change for the selected user account

$LastPWtxt = New-Object System.Windows.Forms.Label

$LastPWtxt.Text = "Last Password Set:"

$LastPWtxt.Location = New-Object System.Drawing.Point(5, 40)

$LastPWtxt.AutoSize = $true

$main_form.Controls.Add($LastPWtxt)


$LastPWtime = New-Object System.Windows.Forms.Label

$LastPWtime.Text = ""

$LastPWtime.Location = New-Object System.Drawing.Point(110, 40)

$LastPWtime.AutoSize = $true

$main_form.Controls.Add($LastPWtime)

# ----------------------------------------------------------------------
# AD Domain Controllers Label
$ADDomainControllersLabel = New-Object System.Windows.Forms.Label

$ADDomainControllersLabel.Text = "Current AD Domain Controller :"

$ADDomainControllersLabel.Location = New-Object System.Drawing.Point(290, 140)

$ADDomainControllersLabel.AutoSize = $true

$main_form.Controls.Add($ADDomainControllersLabel)

# ----------------------------------------------------------------------
# AD Domain Controllers Label
$ADDomainControllersText = New-Object System.Windows.Forms.Label

$ADDomainControllersText.Text = (Get-ADDomainController | Select-Object HostName -ExpandProperty HostName)

$ADDomainControllersText.Location = New-Object System.Drawing.Point(290, 160)

$ADDomainControllersText.AutoSize = $true

$main_form.Controls.Add($ADDomainControllersText)

# ----------------------------------------------------------------------
# Account Detail Lable
$AccountDetailLabel = New-Object System.Windows.Forms.Label

$AccountDetailLabel.Text = "Account Detail :"

$AccountDetailLabel.Location = New-Object System.Drawing.Point(5, 100)

$AccountDetailLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailLabel)
# ----------------------------------------------------------------------
# Account Name Lable
$AccountDetailNameLabel = New-Object System.Windows.Forms.Label

$AccountDetailNameLabel.Text = "Name: "

$AccountDetailNameLabel.Location = New-Object System.Drawing.Point(5, 120)

$AccountDetailNameLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailNameLabel)
# Account Name Display Box
$AccountDetailNameText = New-Object System.Windows.Forms.TextBox

$AccountDetailNameText.Text = ""

$AccountDetailNameText.Location = New-Object System.Drawing.Point(80, 120)

$AccountDetailNameText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailNameText)
# ----------------------------------------------------------------------
# Account User Name Lable
$AccountDetailUserNameLabel = New-Object System.Windows.Forms.Label

$AccountDetailUserNameLabel.Text = "User Name: "

$AccountDetailUserNameLabel.Location = New-Object System.Drawing.Point(5, 140)

$AccountDetailUserNameLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailUserNameLabel)
# Account User Name Display Box
$AccountDetailUserNameText = New-Object System.Windows.Forms.TextBox

$AccountDetailUserNameText.Text = ""

$AccountDetailUserNameText.Location = New-Object System.Drawing.Point(80, 140)

$AccountDetailUserNameText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailUserNameText)
# ----------------------------------------------------------------------
# Account Workstation Name Label

$AD_ComputerNameLabel = New-Object System.Windows.Forms.Label

$AD_ComputerNameLabel.Text = "Workstation: "

$AD_ComputerNameLabel.Location = New-Object System.Drawing.Point(5, 160)

$AD_ComputerNameLabel.AutoSize = $true

$main_form.Controls.Add($AD_ComputerNameLabel)
# Account Workstation Display Box

$AD_ComputerNameText = New-Object System.Windows.Forms.TextBox

$AD_ComputerNameText.Text = ""

$AD_ComputerNameText.Location = New-Object System.Drawing.Point(80, 160)

$AD_ComputerNameText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AD_ComputerNameText)
#----------------------------------------------------------------------
# User Department Name Lable
$AccountDetailDepartmentLable = New-Object System.Windows.Forms.Label

$AccountDetailDepartmentLable.Text = "Department: "

$AccountDetailDepartmentLable.Location = New-Object System.Drawing.Point(5, 180)

$AccountDetailDepartmentLable.AutoSize = $true

$main_form.Controls.Add($AccountDetailDepartmentLable)
# User Department Name Display Box
$AccountDetailDepartmentText = New-Object System.Windows.Forms.TextBox

$AccountDetailDepartmentText.Text = ""

$AccountDetailDepartmentText.Location = New-Object System.Drawing.Point(80, 180)

$AccountDetailDepartmentText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailDepartmentText)
#----------------------------------------------------------------------
# User Office Name Lable
$AccountDetailUserOfficeLabel = New-Object System.Windows.Forms.Label

$AccountDetailUserOfficeLabel.Text = "Office: "

$AccountDetailUserOfficeLabel.Location = New-Object System.Drawing.Point(5, 200)

$AccountDetailUserOfficeLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailUserOfficeLabel)
# User Office Name Display Box
$AccountDetailUserOfficeText = New-Object System.Windows.Forms.TextBox

$AccountDetailUserOfficeText.Text = ""

$AccountDetailUserOfficeText.Location = New-Object System.Drawing.Point(80, 200)

$AccountDetailUserOfficeText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailUserOfficeText)

#----------------------------------------------------------------------
# User Office Phone Lable
$AccountDetailOfficePhoneLabel = New-Object System.Windows.Forms.Label

$AccountDetailOfficePhoneLabel.Text = "Office Phone: "

$AccountDetailOfficePhoneLabel.Location = New-Object System.Drawing.Point(5, 220)

$AccountDetailOfficePhoneLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailOfficePhoneLabel)
# User Office Name Display Box
$AccountDetailOfficePhoneText = New-Object System.Windows.Forms.TextBox

$AccountDetailOfficePhoneText.Text = ""

$AccountDetailOfficePhoneText.Location = New-Object System.Drawing.Point(80, 220)

$AccountDetailOfficePhoneText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailOfficePhoneText)

#----------------------------------------------------------------------
# User Email Address Lable
$AccountDetailEmailLabel = New-Object System.Windows.Forms.Label

$AccountDetailEmailLabel.Text = "Email: "

$AccountDetailEmailLabel.Location = New-Object System.Drawing.Point(5, 240)

$AccountDetailEmailLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailEmailLabel)
# User Email Address Display Box
$AccountDetailEmailText = New-Object System.Windows.Forms.TextBox

$AccountDetailEmailText.Text = ""

$AccountDetailEmailText.Location = New-Object System.Drawing.Point(80, 240)

$AccountDetailEmailText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailEmailText)
#----------------------------------------------------------------------
# User Description Lable
$AccountDetailDescriptionLabel = New-Object System.Windows.Forms.Label

$AccountDetailDescriptionLabel.Text = "Description: "

$AccountDetailDescriptionLabel.Location = New-Object System.Drawing.Point(5, 260)

$AccountDetailDescriptionLabel.AutoSize = $true

$main_form.Controls.Add($AccountDetailDescriptionLabel)
# User Description Display Box
$AccountDetailDescriptionText = New-Object System.Windows.Forms.TextBox

$AccountDetailDescriptionText.Text = ""

$AccountDetailDescriptionText.Location = New-Object System.Drawing.Point(80, 260)

$AccountDetailDescriptionText.Size = New-Object System.Drawing.Size(200, 20)

$main_form.Controls.Add($AccountDetailDescriptionText)
# ----------------------------------------------------------------------
# AD Group Lable
$AccountMemberOf = New-Object System.Windows.Forms.Label

$AccountMemberOf.Text = "AccountMemberOf: "

$AccountMemberOf.Location = New-Object System.Drawing.Point(5, 290)

$AccountMemberOf.AutoSize = $true

$main_form.Controls.Add($AccountMemberOf)
# AD Group Detail
$AccountMemberOfDetail = New-Object System.Windows.Forms.label

$AccountMemberOfDetail.Text = " "

$AccountMemberOfDetail.Location = New-Object System.Drawing.Point(5, 310)

$AccountMemberOfDetail.AutoSize = $true

$main_form.Controls.Add($AccountMemberOfDetail)


#----------------------------------------------------------------------
#----------------------------------------------------------------------
#----------------------------------------------------------------------
#Botton Section

$UnlockUserBotton = New-Object System.Windows.Forms.Button

$UnlockUserBotton.Location = New-Object System.Drawing.Size(450, 70)

$UnlockUserBotton.Size = New-Object System.Drawing.Size(120, 23)

$UnlockUserBotton.Text = "Unlock Account"

$main_form.Controls.Add($UnlockUserBotton)

$UnlockUserBotton.Add_Click(

    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
       
(Unlock-ADAccount -Identity $LockoutUsersBox.SelectedItem)
(Unlock-ADAccount -Identity $AccountnameToUserName)

($LockoutUsersBox.Items.Clear())
($LockoutUsersList = Search-ADAccount -LockedOut)

        Foreach ($User in $LockoutUsersList) {

            $LockoutUsersBox.Items.Add($User.SamAccountName);
        }

( $UnlockADuserMessage.popup( "AD account has been unlocked", 0, "AD Account Unlock", 64) )
    }

)

# ----------------------------------------------------------------------

$UpdateLockUserBotton = New-Object System.Windows.Forms.Button

$UpdateLockUserBotton.Location = New-Object System.Drawing.Size(320, 70)

$UpdateLockUserBotton.Size = New-Object System.Drawing.Size(120, 23)

$UpdateLockUserBotton.Text = "Update Lockout List"

$main_form.Controls.Add($UpdateLockUserBotton)

$UpdateLockUserBotton.Add_Click(
    {
($LockoutUsersBox.Items.Clear())
($LockoutUsersList = Search-ADAccount -LockedOut)
        Foreach ($User in $LockoutUsersList) {
            [void]$LockoutUsersBox.Items.Add($User.SamAccountName);
        }
    }
)

$RDPUserBotton = New-Object System.Windows.Forms.Button

$RDPUserBotton.Location = New-Object System.Drawing.Size(450, 10)

$RDPUserBotton.Size = New-Object System.Drawing.Size(120, 23)

$RDPUserBotton.Text = "RDP to Select User"

$main_form.Controls.Add($RDPUserBotton)

$RDPUserBotton.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        mstsc /v: (Get-ADUser -identity $AccountnameToUserName -Properties otherLoginWorkstations | Select-Object otherLoginWorkstations -ExpandProperty otherLoginWorkstations)
    }

)

$CheckAccountBotton = New-Object System.Windows.Forms.Button

$CheckAccountBotton.Location = New-Object System.Drawing.Size(320, 10)

$CheckAccountBotton.Size = New-Object System.Drawing.Size(120, 23)

$CheckAccountBotton.Text = "Check Account Detail"

$main_form.Controls.Add($CheckAccountBotton)

$CheckAccountBotton.Add_Click(

    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        # Commands
        $LastPWtime.Text = [datetime]::FromFileTime((Get-ADUser -identity $AccountnameToUserName -Properties pwdLastSet).pwdLastSet).ToString('MM/dd/yyyy  hh:mm:ss')
        $AccountDetailNameText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object Name -ExpandProperty Name)
        $AccountDetailUserNameText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object samaccountname -ExpandProperty samaccountname)
        $AD_ComputerNameText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties otherLoginWorkstations | Select-Object otherLoginWorkstations -ExpandProperty otherLoginWorkstations)
        $AccountDetailDepartmentText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object Department -ExpandProperty Department)
        $AccountDetailUserOfficeText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object Office -ExpandProperty Office)
        $AccountDetailOfficePhoneText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object OfficePhone -ExpandProperty OfficePhone)
        $AccountDetailEmailText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object EmailAddress -ExpandProperty EmailAddress)
        $AccountDetailDescriptionText.Text = (Get-ADUser -identity $AccountnameToUserName -Properties * | Select-object Description -ExpandProperty Description)
        $AccountMemberOfDetail.Text = (Get-aduser -identity $AccountnameToUserName -properties MemberOf).memberof | Get-ADgroup | select-object name -ExpandProperty Name | ForEach-Object { $_ + "`n" }
    }

)

$AddADGroup = New-Object System.Windows.Forms.Button

$AddADGroup.Location = New-Object System.Drawing.Size(320, 320)

$AddADGroup.Size = New-Object System.Drawing.Size(120, 23)

$AddADGroup.Text = "Add Group"

$main_form.Controls.Add($AddADGroup)

$AddADGroup.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        
        $result = $ADDADGroupConfirm.Popup("Are you want to add this group to $AccountnameToUserName ?", 5, "Add Group Confirmation", 48 + 4)
        if ($result -eq 6) {
            Add-ADGroupMember -Identity $ADGroupListBox.SelectedItem -Members $AccountnameToUserName -Confirm:$false
        
            if ($result -eq 7) { $AccountMemberOfDetail.Text = (Get-aduser -identity $AccountnameToUserName -properties MemberOf).memberof | Get-ADgroup | select-object name -ExpandProperty Name | ForEach-Object { $_ + "`n" } }


            $AccountMemberOfDetail.Text = (Get-aduser -identity $AccountnameToUserName -properties MemberOf).memberof | Get-ADgroup | select-object name -ExpandProperty Name | ForEach-Object { $_ + "`n" }
        }


    }

)
$RemoveADGroup = New-Object System.Windows.Forms.Button

$RemoveADGroup.Location = New-Object System.Drawing.Size(450, 320)

$RemoveADGroup.Size = New-Object System.Drawing.Size(120, 23)

$RemoveADGroup.Text = "Remove Group"

$main_form.Controls.Add($RemoveADGroup)

$RemoveADGroup.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        
        $result = $RemoveADGroupConfirm.Popup("Are you want to Remove this group for $AccountnameToUserName ?", 5, "Remove Group Confirmation", 48 + 4)
        if ($result -eq 6) { Remove-ADGroupMember -Identity $ADGroupListBox.SelectedItem -Members $AccountnameToUserName -Confirm:$false }    

        
        if ($result -eq 7) { $AccountMemberOfDetail.Text = (Get-aduser -identity $AccountnameToUserName -properties MemberOf).memberof | Get-ADgroup | select-object name -ExpandProperty Name | ForEach-Object { $_ + "`n" } }


        $AccountMemberOfDetail.Text = (Get-aduser -identity $AccountnameToUserName -properties MemberOf).memberof | Get-ADgroup | select-object name -ExpandProperty Name | ForEach-Object { $_ + "`n" }



    }


)


$ResetADUserPw = New-Object System.Windows.Forms.Button

$ResetADUserPw.Location = New-Object System.Drawing.Size(450, 100)

$ResetADUserPw.Size = New-Object System.Drawing.Size(120, 23)

$ResetADUserPw.Text = "Reset Password"

$main_form.Controls.Add($ResetADUserPw)

$ResetADUserPw.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        $result = $ResetADUserPwConfirm.Popup("Are you want to Reset Password for $AccountnameToUserName ?", 5, "Reset Password Confirmation", 48 + 4)
        if ($result -eq 6) {      
            $defaultpw = ConvertTo-SecureString -AsPlainText "ChangePassword" -Force
            Set-ADAccountPassword -Identity $AccountnameToUserName -NewPassword $defaultpw -Reset
            Set-ADUser -Identity $AccountnameToUserName -ChangePasswordAtLogon $true       
        }    
        
        if ($result -eq 7) { }
    }

)

$ClearAllGroupFromUser = New-Object System.Windows.Forms.Button

$ClearAllGroupFromUser.Location = New-Object System.Drawing.Size(450, 360)

$ClearAllGroupFromUser.Size = New-Object System.Drawing.Size(120, 23)

$ClearAllGroupFromUser.Text = "Remove ALL Group"

$main_form.Controls.Add($ClearAllGroupFromUser)

$ClearAllGroupFromUser.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        $result = $ClearAllGroupFromUserConfirm.Popup("Are you want to Remove All Group for $AccountnameToUserName ?", 5, "Remove All Group Confirmation", 48 + 4)
        if ($result -eq 6) {      
            $Groups = (Get-ADUser -Identity $AccountnameToUserName -Properties memberOf).memberOf
            ForEach ($Group In $Groups) {
                Remove-ADGroupMember -Identity $Group -Members $AccountnameToUserName -Confirm:$false
            }
        }    
        
        if ($result -eq 7) { }
    }

)

$DisableUserAccount = New-Object System.Windows.Forms.Button

$DisableUserAccount.Location = New-Object System.Drawing.Size(320, 360)

$DisableUserAccount.Size = New-Object System.Drawing.Size(120, 23)

$DisableUserAccount.Text = "Disable Account"

$main_form.Controls.Add($DisableUserAccount)

$DisableUserAccount.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        $result = $DisableUserAccountConfirm.Popup("Are you want to Remove All Group for $AccountnameToUserName ?", 5, "Disable Account Confirmation", 48 + 4)
        if ($result -eq 6) {      
            Disable-ADAccount -Identity $AccountnameToUserName -Confirm:$false
        }        
        if ($result -eq 7) { }
    }

)

$TerminateUserAccount = New-Object System.Windows.Forms.Button

$TerminateUserAccount.Location = New-Object System.Drawing.Size(450, 400)

$TerminateUserAccount.Size = New-Object System.Drawing.Size(120, 23)

$TerminateUserAccount.Text = "Terminate Account"

$main_form.Controls.Add($TerminateUserAccount)

$TerminateUserAccount.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        $result = $TerminateUserAccountConfirm.Popup("Do you want to Terminate User Account for $AccountnameToUserName ?", 5, "Terminate Account Confirmation", 48 + 4)
        if ($result -eq 6) {
            $Groups = (Get-ADUser -Identity $AccountnameToUserName -Properties memberOf).memberOf
            ForEach ($Group In $Groups) {
                Remove-ADGroupMember -Identity $Group -Members $AccountnameToUserName -Confirm:$false
            }      
            Disable-ADAccount -Identity $AccountnameToUserName -Confirm:$false
        
            $DistinguishedNameToOU = Get-aduser -identity $AccountnameToUserName -properties * | Select-Object DistinguishedName -ExpandProperty DistinguishedName
            $TerminatedUsersOU = "OU=TerminatedUsers,OU=IT,DC=***,DC=***"
            Move-ADObject -Identity "$DistinguishedNameToOU" -TargetPath $TerminatedUsersOU
        
        
        }        
        if ($result -eq 7) { }
    }

)

$TerminateComputerAccount = New-Object System.Windows.Forms.Button

$TerminateComputerAccount.Location = New-Object System.Drawing.Size(320, 400)

$TerminateComputerAccount.Size = New-Object System.Drawing.Size(120, 23)

$TerminateComputerAccount.Text = "Terminate AD PC"

$main_form.Controls.Add($TerminateComputerAccount)

$TerminateComputerAccount.Add_Click(
    {
        # Convert Account Name to User Name
        $AccountnameToUserName = $ADuserBox.selectedItem
        $AccountnameToUserName = (Get-ADUser -Filter 'Name -like $AccountnameToUserName' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname)
        $result = $TerminateComputerAccountConfirm.Popup("Do you want to Terminate User Account for $AccountnameToUserName ?", 5, "Terminate Account Confirmation", 48 + 4)
        if ($result -eq 6) {
            

            Set-ADComputer -Identity KLAPORTEN -Enabled $false
            $DistinguishedNameToOU = Get-ADComputer (Get-ADUser -identity $AccountnameToUserName -Properties otherLoginWorkstations | Select-Object otherLoginWorkstations -ExpandProperty otherLoginWorkstations) | Select-Object DistinguishedName -ExpandProperty DistinguishedName
            $TerminatedComputerOU = "OU=TerminatedComputers,OU=IT,DC=***,DC=***"
            Move-ADObject -Identity "$DistinguishedNameToOU" -TargetPath $TerminatedComputerOU                                                                                                                                                            

        }        
        if ($result -eq 7) { }
    }

)


$LoadAllAccount = New-Object System.Windows.Forms.Button

$LoadAllAccount.Location = New-Object System.Drawing.Size(230, 10)

$LoadAllAccount.Size = New-Object System.Drawing.Size(80, 23)

$LoadAllAccount.Text = "Active User"

$main_form.Controls.Add($LoadAllAccount)

$LoadAllAccount.Add_Click(
    {
        # Convert Account Name to User Name
        $Users = (Get-ADUser -Filter * -Property Enabled | Where-Object { $_.Enabled -like "True" } | Select-Object Name -ExpandProperty Name | Sort-Object Name)
        ($ADuserBox.Items.Clear())
        Foreach ($User in $Users) {
        
            [void]$ADuserBox.Items.Add($User.Name);
        
        }
    }

)

$ShowDisableAccount = New-Object System.Windows.Forms.Button

$ShowDisableAccount.Location = New-Object System.Drawing.Size(320, 40)

$ShowDisableAccount.Size = New-Object System.Drawing.Size(120, 23)

$ShowDisableAccount.Text = "ShowDisableAccount"

$main_form.Controls.Add($ShowDisableAccount)

$ShowDisableAccount.Add_Click(
    {
        # Convert Account Name to User Name
        $Users = (Get-ADUser -Filter * -Property Enabled | Where-Object { $_.Enabled -like "False" } | Select-Object Name -ExpandProperty Name | Sort-Object Name)
        ($ADuserBox.Items.Clear())
        Foreach ($User in $Users) {
        
            [void]$ADuserBox.Items.Add($User.Name);
        
        }
    }

)

$ShowAllAccount = New-Object System.Windows.Forms.Button

$ShowAllAccount.Location = New-Object System.Drawing.Size(450, 40)

$ShowAllAccount.Size = New-Object System.Drawing.Size(120, 23)

$ShowAllAccount.Text = "ShowAllAccount"

$main_form.Controls.Add($ShowAllAccount)

$ShowAllAccount.Add_Click(
    {
        # Convert Account Name to User Name
        $Users = (get-aduser -filter * -Properties SamAccountName)
        ($ADuserBox.Items.Clear())
        Foreach ($User in $Users) {
        
            [void]$ADuserBox.Items.Add($User.Name);
        
        }
    }

)

#End of main window
$main_form.ShowDialog() | Out-Null
