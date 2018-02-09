# Credit goes to: Donia Gamal ElDin (donia.gamaleldin@gmail.com )

# GoogleDriveFolderSharing

This is a simple tutorial that uses Google Dive Api V2 to create folders and share them with specific email addresses.

To make it work [First time only]:
1- Go to https://console.developers.google.com/ 
2- Go to Credentials tab
3- Create a new credential
4- Then download the json file


1- Create an Excel file with 3 columns, the first is the Folder Name and the second is the email to share the folder with, the third is the department [if exists].
2- Put the json file beside the exe file. 
3- Open the config file and replace the following:
	a- FileName : The full path of the Groups excel file
	b- RootFolderName : The folder that the folders should be created underneath
	c- AuthFilePath : The path of the auth file (the one downloaded in 4)
	d- SharingMessage : The message that will be sent when updating sharing