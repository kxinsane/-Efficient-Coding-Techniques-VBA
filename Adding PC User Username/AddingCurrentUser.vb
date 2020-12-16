	'Adding current user of a PC into the Pathname to Folder on local or remote drive
	"C:\Users\" & Environ$("Username") & "\Folder One\Folder Two\Folder Three\" & Filename
	
	'main part what we need is Environ$("Username")