#Get Current User Profile
$CurrentUserSID = Get-RegistryKey -Key "HKEY_CURRENT_USER\Volatile Environment" -sid $CurrentConsoleUserSession.sid
$ProfilePath = $CurrentUserSID.USERPROFILE

#Set needed Variables
$OneNotePreMoveExportFile = "$envLocalAppData\Company\OnenotePreMoveFiles.txt"
$OneNotePostMoveExportFile = "$envLocalAppData\Company\OnenotePostMoveFiles.txt"
$NewOneNoteParentPath = "$profilepath\OneNote Notebooks"

#Clear existing Reg Keys
Remove-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Company\OneNoteBooks" -Recurse -ContinueOnError $True

#Get Onenote List of filenames and Paths using .net calls
[void][reflection.assembly]::LoadWithPartialName("Microsoft.Office.Interop.Onenote")
$OneNote = New-Object Microsoft.Office.Interop.Onenote.ApplicationClass

#Get Hierarchy
[Xml]$Xml = $Null
$OneNote.GetHierarchy($Null, [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsNotebooks, [ref] $Xml)

#Create Array
$NotebookPaths = @()

#Search and format Notebooks into an array object
ForEach($Notebook in ($Xml.Notebooks.Notebook)) {
  $arrayObject = New-Object PSObject -Property @{
    'Notebook Names' = $Notebook.name
    'Notebook Location' = $Notebook.path
  }
  $NotebookPaths += $arrayObject 

}
# Pre-Move Notebook locations for debuging purposes
$OneNoteOutput = $Notebookpaths | ft 
$OneNoteOutput  | Out-String -Width 2000 | Out-File $OneNotePreMoveExportFile

#Modify Array items for Notebook Post-Move Locations
foreach($Item in $NotebookPaths){
    if(($Item.'Notebook Location') -like "*C:\users\$env:USERNAME\*"){ 
    $ItemPath = $Item.'Notebook Location'
    $split = Split-Path -Path $itempath -Parent
    $FixedOneNotePath = $itemPath.Replace($split,$NewOneNoteParentPath)
    $Item.'Notebook Location' = $FixedOneNotePath
    $item
    }
}

# Post-Move Notebook locations for The template File
$OneNoteOutput = $Notebookpaths | ft 
$OneNoteOutput  | Out-String -Width 2000 | Out-File $OneNotePostMoveExportFile

#Stop OneNote Process to kill Com Objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)

#Output to Registry
New-Item –Path "HKLM:\SOFTWARE\Company" –Name "OneNoteBooks"
ForEach($Item in $NotebookPaths) {
          $ItemName = $Item.'Notebook Names'
          $ItemPath = $Item.'Notebook Location'
          New-Item –Path "HKLM:\SOFTWARE\Company\OneNoteBooks" –Name "$itemName" -Force
          New-ItemProperty -Path "HKLM:\SOFTWARE\Company\OneNoteBooks\$ItemName" -Name "User" -Value ”$env:USERNAME”  -PropertyType "String" -Force
          New-ItemProperty -Path "HKLM:\SOFTWARE\Company\OneNoteBooks\$ItemName" -Name "Name" -Value ”$ItemName”  -PropertyType "String" -Force
          New-ItemProperty -Path "HKLM:\SOFTWARE\Company\OneNoteBooks\$ItemName" -Name "Location" -Value ”$ItemPath”  -PropertyType "String" -Force
    }
