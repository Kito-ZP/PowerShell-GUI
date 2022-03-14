# PowerShell-GUI



Powershell Tools for maintenance Active Directory<br/>
1. AD Infor Tool V1.1<br/>
-AD Info Tool.ps1 is a Powershell Scrpit runing powershell as GUI tool.<br/>
-Disaplay User info<br/>
-One Click RDP to Select User<br/>
-Terminate User<br/>
-Managing AD groups<br/>
In order to run the script, there are some place need to be edit
I. Line 563 Edit "ChangePassword" to any password you want to setup as default password
After reset password, will turn on ChangePasswordAtLogon feature.
II. Line 649 need to edit the DN for Moving Terminated User to specific container.
III. Line 680 need to edit the DN for Moving Terminated PC to specific container.
