# CopySpotlight
Copies Windows Spotlight Images to your personal Pictures folder

## Install
See: https://docs.microsoft.com/en-us/dotnet/framework/windows-services/how-to-install-and-uninstall-services

## MS Teams Background Configuration
To have spotlight images available as MS Teams background, you have to upload one custom background in MS Teams once.

You should find the uploaded image in  
C:\Users\YOUR_USERNAME\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads.

Copy the name of the uploaded custom image and add it to the config.json file in this service's bin folder.  
``{"filename": "39647c30-576f-4f33-9a69-4ec4b87fed03" }``

This allows CopySpotlight to copy the latest spotlight image to the MS Teams background folder and have it available 
as a background in MS Teams.  
It is always up to date with the latest spotlight image, so you don't have to upload a new custom background in 
MS Teams every time a new spotlight image is available.
