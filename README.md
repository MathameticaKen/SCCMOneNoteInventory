# SCCMOneNoteInventory
Inventory Users Recent OneNote Notebooks via SCCM Inventory  

1.Create the Mof(I used RegKey2MOF)
2.Import the Mof to Hardware inventory classes into client settings
3.Run the PS1 script on the systems you want to inventory

## Note:
There are Functions used from PSApplicationDeployment Toolkit used here.   
This was primary used to assist with OneDrive Known Folder redirection by moving files out of Desktop,documents,Pictures and writing where the New location would be after deployment.   
Read the code - Run at your own Risk.
