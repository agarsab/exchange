#
# PowerShell script to collect newest received and sent item in a user mailbox
# By default, it tries to find the newest message in "Inbox" and "SentItems" folders
# If no message is found, it uses the creation date of the folder itself
#
# Tested on Microsoft Exchange Server 2010
# Created by agarsab@gmail.com
# Get latest update at https://github.com/agarsab/exchange

Function Get-MailboxActivity
{ 
 
[cmdletBinding()] 
 
Param( 
[Parameter(Mandatory=$true)][string]$Identity
) 
 
Process 
{

$Today=Get-Date 

$ReceivedFolderStatistics=Get-MailboxFolderStatistics -identity $Identity -IncludeOldestAndNewestItems -folderscope inbox | where {$_.FolderType -eq "Inbox"} |  select-object Identity,Date,NewestItemReceivedDate,FolderSize,ItemsInFolder

$SentFolderStatistics=Get-MailboxFolderStatistics -identity $Identity -IncludeOldestAndNewestItems -folderscope SentItems | select-object Identity,Name,Date,NewestItemReceivedDate,FolderSize,ItemsInFolder 

$FolderStatistics = New-Object PSObject

add-member -input $FolderStatistics -membertype noteproperty -name "Identity" -value $Identity

if ($ReceivedFolderStatistics.Identity)
{
   #add-member -input $FolderStatistics -membertype noteproperty -name "ReceivedFolderIdentity" -value $ReceivedFolderStatistics.Identity
   add-member -input $FolderStatistics -membertype noteproperty -name "ReceivedFolderDate" -value $ReceivedFolderStatistics.Date
   add-member -input $FolderStatistics -membertype noteproperty -name "ReceivedFolderItems" -value $ReceivedFolderStatistics.ItemsInFolder
   add-member -input $FolderStatistics -membertype noteproperty -name "ReceivedFolderSize" -value $ReceivedFolderStatistics.FolderSize
   add-member -input $FolderStatistics -membertype noteproperty -name "NewestItemReceivedDate" -value $ReceivedFolderStatistics.NewestItemReceivedDate
   if ($ReceivedFolderStatistics.NewestItemReceivedDate)
      {
      add-member -input $FolderStatistics -membertype noteproperty -name "NewestItemReceivedDays" -value ($Today.subtract($ReceivedFolderStatistics.NewestItemReceivedDate).days)
      }
   else
      {
      add-member -input $FolderStatistics -membertype noteproperty -name "NewestItemReceivedDays" -value ($Today.subtract($ReceivedFolderStatistics.Date).days)
      }
}

if ($SentFolderStatistics.Identity)
{
   #add-member -input $FolderStatistics -membertype noteproperty -name "SentFolderIdentity" -value $SentFolderStatistics.Identity
   add-member -input $FolderStatistics -membertype noteproperty -name "SentFolderDate" -value $SentFolderStatistics.Date
   add-member -input $FolderStatistics -membertype noteproperty -name "SentFolderItems" -value $SentFolderStatistics.ItemsInFolder
   add-member -input $FolderStatistics -membertype noteproperty -name "SentFolderSize" -value $SentFolderStatistics.FolderSize
   add-member -input $FolderStatistics -membertype noteproperty -name "NewestItemSentDate" -value $SentFolderStatistics.NewestItemReceivedDate
   if ($SentFolderStatistics.NewestItemReceivedDate) 
      {
      add-member -input $FolderStatistics -membertype noteproperty -name "NewestItemSentDays" -value ($Today.subtract($SentFolderStatistics.NewestItemReceivedDate).days)
      }
   else
      {
      add-member -input $FolderStatistics -membertype noteproperty -name "NewestItemSentDays" -value ($Today.subtract($SentFolderStatistics.Date).days)
      }
}

return $FolderStatistics
}
}
