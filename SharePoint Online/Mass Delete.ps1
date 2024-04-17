#Config Variables
$SiteURL = "https://mhud.sharepoint.com/sites/infomgmt"
$ListName ="Infosec"
 
 
Try {
    #Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive
     
    #Get All List Items in Batch
    $ListItems = Get-PnPListItem -List $ListName -PageSize 1000 | Sort-Object ID -Descending
 
    #sharepoint online powershell delete all items in a list
    ForEach ($Item in $ListItems)
    {
        Remove-PnPListItem -List $ListName -Identity $Item.Id -Force
    }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}


https://mhud.sharepoint.com/:f:/r/sites/infomgmt/infosec/Audit/ITGC%20audit%202023/Evidence/6%20Network/Firewall%20and%20Network%20Diagrams?csf=1&web=1&e=1NWlhp

# Specify the library and folder path
$libraryName = "https://mhud.sharepoint.com"
$folderRelativeUrl = "/sites/infomgmt/infosec/Audit/ITGC audit 2023/Evidence/6 Network/Firewall and Network Diagrams"

# Get the folder item
$folderItem = Get-PnPFolder -Url $folderRelativeUrl -ErrorAction Stop

# Now you can interact with $folderItem, or get files/items within it
Get-PnPListItem 

# Get all items within the folder
$items = Get-PnPListItem -FolderServerRelativeUrl $folderRelativeUrl -PageSize 500 -List infosec

# Loop through each item and display its title
foreach ($item in $items) {
    Write-Host "Item Title: $($item.FieldValues['FileLeafRef'])"
}


$pageSize = 1000  # Number of items to retrieve per page
$position = $null  # Position to start retrieving items

do {
    $items = Get-PnPListItem -FolderServerRelativeUrl $folderRelativeUrl -PageSize $pageSize
    # Process items
    foreach ($item in $items) {
        Write-Host "Item Title: $($item.FieldValues['FileLeafRef'])"
    }

    # Get the position of the next batch of items
    $position = $items.ListItemCollectionPosition
}
while ($position -ne $null)


Get-PnPList -Identity $ListName -ThrowExceptionIfListNotFound | Get-PnPListItem -FolderServerRelativeUrl $folderRelativeUrl -PageSize 100 -ScriptBlock { 
    Param($items) Invoke-PnPQuery } | ForEach-Object {$_.Recycle()
 }


 #Get List Items to Delete
$ListItems =  Get-PnPListItem -List $ListName -PageSize 500
 
#Create a New Batch
$Batch = New-PnPBatch
 
#Clear All Items in the List
ForEach($Item in $ListItems)
{    
     Remove-PnPListItem -List $ListName -Identity $Item.ID -Recycle -Batch $Batch
}
 
#Send Batch to the server
Invoke-PnPBatch -Batch $Batch


1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
$SiteURL = "https://crescent.sharepoint.com/sites/Operations"
$ListName = "Inventory"
 
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
 
#Get List Items to Delete
$ListItems =  Get-PnPListItem -List $ListName -FolderServerRelativeUrl $folderRelativeUrl -PageSize 500
 
#Create a New Batch
$Batch = New-PnPBatch
 
#Clear All Items in the List
ForEach($Item in $ListItems)
{    
     Remove-PnPListItem -List $ListName -Identity $Item.ID -Recycle -Batch $Batch
}
 
#Send Batch to the server
Invoke-PnPBatch -Batch $Batch


#Get all list items from list in batches  
$ListItems = Get-PnPFolderItem -FolderSiteRelativeUrl $folderRelativeUrl  
   
Write-host "Total Number of List Items:" $($ListItems.Count)  
   
#Loop through each Item  
ForEach($Item in $ListItems)  
{   
    Write-Host "Id :" $Item["ID"]  
    Write-Host "Title :" $Item["Title"]  
}  