/********
This is a script that parses through downloaded files in a subfolder,
and detects what type of video game they contain, moving detected game
types to a different subfolder for the emulator to run.
It is a work in progress: Games that are present as disc images need to
be mounted as virtual drives, so that their internal files can be
examined, and I haven't implemented that logic yet.
*********/

//constants used by script logic
var FOLDER_DELIM = '\\';
var ZIP_EXECUTABLE_PATH = '"C:\\Users\\stile\\OneDrive\\Documents\\!Apps\\7-ZipPortable\\App\\7-Zip\\7z.exe"'; //change this to wherever you keep your 7zip

//load up global variables
var fso = new ActiveXObject('Scripting.FileSystemObject');
var shell = new ActiveXObject('Wscript.Shell')
var baseFolderName = fso.GetParentFolderName(WScript.ScriptFullName) + FOLDER_DELIM;
var incomingFolder = getOrMakeFolder(baseFolderName + '!INCOMING');
var romsFolder = getOrMakeFolder(baseFolderName + 'BAT_SHARE' + FOLDER_DELIM + 'roms');
var totalFilesMoved = 0;
var totalBytesMoved = 0;

//actual script logic
for (var files = new Enumerator(incomingFolder.files); !files.atEnd(); files.moveNext()){
 importFile(files.item());
}
log("Moved " + totalBytesMoved + " bytes in " + totalFilesMoved + " files.");

//folder/file processing functions

function detectType(fileOrFolder){
	if (fileOrFolder.Files){ //if we are a folder, we need to iterate through all the files (including subdirectories) to see if any are identified
		for (var files = new Enumerator(fileOrFolder.files); !files.atEnd(); files.moveNext()){
			var result = detectType(files.item());
			if (result){return result;}
  }
		for (var folders = new Enumerator(fileOrFolder.folders); !folders.atEnd(); folders.moveNext()){
			var result = detectType(folders.item());
			if (result){return result;}
  }
		return null; //if none of the files in the directory helped, then return null to indicate type could NOT be identified
	} else {
		switch (getExtension(fileOrFolder.name)){ //if we are a file, then handle differently depending on extension (some can be identified easily, some require more checking)
		 case 'sfc':    return 'snes';	
			case 'gcz':    return 'gamecube';
			case 'gg':     return 'gamegear';	
			case 'gb':     return 'gb';
			case 'gba':    return 'gba';
			case 'gbc':    return 'gbc';
			case 'sms':    return 'mastersystem';
			case 'md':     return 'megadrive';
			case 'nds':    return 'nds';
			case 'a26':    return 'atari2600';
			case 'a52':    return 'atari5200';
			case 'a78':    return 'atari7800';
			case 'pce':    return 'pcengine';
			case 'pygame': return 'pygame';
			case 'd64':	   return 'c64';
			case 'z64':    return 'n64';
			case 'fds': case 'nes': 
			 return 'nes';
		 //TODO: learn how to recognize more file types!
   default:
		  return null;		
	 }
	}
}

function extractAllFrom(file){
	var cmd;
	var destPath = file.path + '_EXTRACTED';
	switch (getExtension(file.name)){
		case 'zip':
		case '7z':
		 cmd = ZIP_EXECUTABLE_PATH + ' x "' + file.path + '" "-o' + destPath + '" -aoa';
			break;
		//TODO: implement	additional types of files that can be extracted from (rar, chd, iso, etc)
		default:
		 cmd = null;
	}
	if (cmd){
		try {
		 shell.run(cmd, 0, true);
		 return fso.GetFolder(destPath);
		} catch (error){
			return file; //if we couldn't extract successfully, simply return the original file
		}
	} else {
		return file;
	}
}

function importFile(file){
	var toCopy = extractAllFrom(file); //if original file is an archive, toMove will be a folder with the contents, otherwise it will just be the original file
	var detectedType = detectType(toCopy);
	if (detectedType){
		var bytesCopied = copyFileOrFolder(toCopy, romsFolder.Path + FOLDER_DELIM + detectedType + FOLDER_DELIM + file.name.split('.')[0] + FOLDER_DELIM)
		if (bytesCopied > 0){
			file.Delete(); //ONLY delete the original file if the copy was SUCCESSFUL
			totalFilesMoved++;
			totalBytesMoved += bytesCopied;
			return bytesCopied;
		} else {
			return 0;
		}
	} else {
		log ('Unknown file type ' + detectedType + ', cannot process: ' + file.name);
		return 0;
	}
}

function getOrMakeFolder(folderName){
	if (fso.FolderExists(folderName)){
		return fso.GetFolder(folderName);
	} else {
		var parentFolderName = fso.GetParentFolderName(folderName);
		if (!fso.FolderExists(parentFolderName)){getOrMakeFolder(parentFolderName);}
		return fso.CreateFolder(folderName);
	}
}

function copyFileOrFolder(fileOrFolder, destPath){ //this function works because both files and folders have Size, Copy(), and Delete()
	log('Copying ' + fileOrFolder.Name + ' to ' + destPath);
	var copySize = fileOrFolder.Size;
	try {
		getOrMakeFolder(destPath);
		var copyFromPath = (fileOrFolder.Files ? fileOrFolder.Path + "\\*.*" : fileOrFolder.Path);
		fso.CopyFile(copyFromPath, destPath, true);
 } catch (error){
		log(error.description);
		return 0;
	}
	return copySize;
}

function getExtension(filepath){
	return filepath.split('.').pop().toLowerCase();
}	

function popup(msg){
	try {
	 log(msg);
	} catch(error){
		msg = msg + " <- (failed to log message)"
	}
	WScript.echo(msg);
}

function fatal(msg){
	popup("FATAL ERROR: " + msg);
	WScript.Quit();
}

function log(msg){
	if (!this.eApp){
	 try {
	  this.eApp = new ActiveXObject("Excel.Application");
		 this.eApp.visible = true;
		} catch (error){ //if the user does not have excel in their system, we cannot create the log, so simply exit the log function
			this.eApp = -1; //flag the static variable so that we know not to keep trying
  }			
	}	
	if (this.eApp == -1){return false;}
	if (!this.logSheet){
		this.logSheet = eApp.Workbooks.Add.Worksheets(1);
		this.outRow = 1;
	}
 this.logSheet.Cells(this.outRow++,1).Value2 = msg;
	if (outRow > 10){this.eApp.ActiveWindow.ScrollRow = outRow - 9;}
	this.logSheet.Parent.Saved = true;
	return true;
}
