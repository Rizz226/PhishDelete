Write-Host "The script finds and deletes emails from user mailboxes in Exchange.

'Compliance Search' and 'Search And Purge' roles are required in MSFT Security and Compliance Center.

A maximum of 10 items per mailbox can be removed at one time." -ForegroundColor Yellow

 

##Importing modules if not already imported

$AlreadyImportedModules = Get-Module
$ModulesToCheck = @("ExchangeOnlineManagement")

ForEach($i in $ModulesToCheck){
If($AlreadyImportedModules.Name -notcontains $i){
Import-Module $i
}
}

##Connecting to Security & Compliance Center
$Username = Read-Host "Please enter your Protegrity email address"
Connect-IPPSSession -UserPrincipalName $Username



##Finding the email

$Name = Read-Host "Enter a search title"
$ExchangeLocation = Read-Host "Specify All to search all mailboxes. To specify multiple mailboxes or distribution groups, enter their email address separated by comma"
$ExchangeLocation2 = $ExchangeLocation.Split(",").Trim()
$ContentMatchQuery = Read-Host "Please specify content search query in the format: (From:email@address.com) AND (Received:12/14/2021..12/15/2021) AND (Subject:'Phishing Email')"
$Name2 = $Name + "_purge"

New-ComplianceSearch -Name $Name -ExchangeLocation $ExchangeLocation2 -ContentMatchQuery $ContentMatchQuery | Out-Null
Start-ComplianceSearch $Name | Out-Null

While((Get-ComplianceSearch $Name).Status -ne "Completed"){
Write-Host "Waiting for 60 seconds for the search to complete" -ForegroundColor Yellow
Start-Sleep -Seconds 60
}

Get-ComplianceSearch $Name | FL Name,Status,ExchangeLocation,PublicFolderLocation,ContentMatchQuery,Items,Errors,NumFailedSources,@{Name="Non0Results";Expression={(Get-ComplianceSearch $Name).SuccessResults -Split "`n" -NotLike "item count: 0"}}

Read-Host "Please VERIFY the search results above. Press Enter to soft delete the email or Ctrl+C to exit"



##Deleting the email
##Change the purge type between Soft and Hard. Hard purge is unrecoverable.
New-ComplianceSearchAction -SearchName $Name -Purge -PurgeType SoftDelete -Confirm:$False | Out-Null

While((Get-ComplianceSearchAction $Name2).Status -ne "Completed"){
Write-Host "Waiting for 60 seconds for the delete to complete
" -ForegroundColor Yellow
Start-Sleep -Seconds 60
}

Write-Host "The final delete action results are as following:" -ForegroundColor Yellow

Get-ComplianceSearchAction $Name2 | FL SearchName,Status,Errors,Results
