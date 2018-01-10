; Where ALL of this is stored: https://www.dropbox.com/sh/rytpiimr5snjnbk/AAB_nORecls7kcCB5ldlV-tAa?dl=0

IniRead,Current_ini,version.ini,Info,Version,N/A  ;This is the desktop version.ini which will be compared with the online version.ini. It assumes that the version.ini is in the same directory unless specified otherwise.
Update_URL := "https://www.dropbox.com/s/keu8jhvgkvajka4/Version.ini?dl=1" ; "https://drive.google.com/uc?export=download&id=1OHgQ_oiUT5bCAAyLix-7LW6Z9QyBRajB" ;The URL of the online version.ini file for your script
	
Random,Filler,10000000,99999999
	, Version_File := A_Temp . "\" . Filler . ".ini"
	, Temp_Exe := A_Temp . "\" . f . "ShortcutToolkit.ahk" ;replace yourfilename.exe with your desired program name. 

UrlDownloadToFile,%Update_URL%,%Version_File%
IniRead,Version,%Version_File%,Info,Version,N/A
If (Version = "N/A"){
			FileDelete,%Version_File%

			
		}
If (Version > Current_ini)
{
		IniRead,URL,%Version_File%,Info,URL,N/A
		UrlDownloadToFile, %url%, %Temp_Exe%
		Run, %Temp_Exe%
		UrlDownloadToFile,%Update_URL%,version.ini ;this makes the newer version.ini replace the current version.ini on your desktop. It assumes that the version.ini is in the same directory unless specified otherwise. Make sure this matches the location of the old version.ini
		exitapp
}
	Else
{
IfExist, %Temp_Exe%
{
	FileDelete,%Version_File%
	Run, %Temp_Exe%
}
IfNotExist, %Temp_Exe%
{	
Msgbox, File not found. Downloading. ;you can remove this line if you don't want the message box.
	UrlDownloadToFile,%Update_URL%,%Version_File%
	IniRead,URL,%Version_File%,Info,URL,N/A
	UrlDownloadToFile, %url%, %Temp_Exe%
	FileDelete,%Version_File%
	Run, %Temp_Exe%
}
exitapp
}

exitapp