'This is a script that synchronizes files between a local folder and 
'a USB stick containing Batocera (a Linux build that runs old video
'games). It loops recursively through specific subfolders in both locations,
'and copies the most recent version of any matching files so that both
'locations have the latest file. This allows me to always keep a 
'backup of all my old game files, as well as configuration files,
'save files, game media, etc

option explicit

'SCRIPT BEHAVIOR CONSTANTS

const DEBUG_MODE = true
const FOLDER_DELIM = "\"

'INITIALIZE LOG

const MAX_LOG_COL_WIDTH = 150 'the maximum width of a column in your log sheet
const SAVE_AFTER_EVERY_LOG = false 'it is NOT recommended to use true here, as it will slow down your script quite a bit...MAY be useful for debugging in some cases

dim startTime: starttime = now
dim LOG_SHEET:    getOrMakeLogSheet(0)

'INITIALIZE VARIABLES

dim totalFilesCopied: totalFilesCopied = 0
dim totalBytesCopied: totalBytesCopied = 0

dim BASE_FOLDER:  set BASE_FOLDER  = getOrMakeFolder(getParentFolderPath(wscript.scriptfullname))

dim BAT_DRIVE:    set BAT_DRIVE    = getDriveFromName(BATOCERA)
dim BAT_FOLDER:   set BAT_FOLDER   = getOrMakeFolder(subfolderpath( BASE_FOLDER, BATOCERA))
if BAT_DRIVE is nothing then fatal "Could not locate BATOCERA thumb drive for syncing!"
syncFile "batocera-boot.conf", bat_drive.rootfolder, bat_folder

dim SHARE_DRIVE:  set SHARE_DRIVE  = getDriveFromName(BAT_SHARE)
dim SHARE_FOLDER: set SHARE_FOLDER = getOrMakeFolder(subfolderpath( BASE_FOLDER, BAT_SHARE))
if SHARE_DRIVE is nothing then fatal "Could not locate BAT_SHARE thumb drive for syncing! If this is a new installation, please boot from the stick, then format the shared drive to exFAT and label it ""BAT_SHARE"" then try again"



dim BLACKLIST_FOLDERS: set BLACKLIST_FOLDERS = createobject("scripting.dictionary")
with BLACKLIST_FOLDERS
 .add ".cache", 0
end with

'ACTUAL SCRIPT LOGIC:

'first, sync batocera config files

syncFolder "system", share_drive.rootfolder, share_folder
'then, sync save files:
syncFolder "saves", share_drive.rootfolder, share_folder
'and the games themselves:
syncFolder "roms", share_drive.rootfolder, share_folder
'close out the script
logAndSave array(totalfilescopied & " files copied at " & transferspeed(totalBytesCopied, now - starttime) & ".", totalbytescopied, timesince(starttime))

'WORKS IN PROGRESS

function logAndSave(msgs)
 log msgs
 log_sheet.parent.save
end function

function transferSpeed(bytes, days)
 dim result
 if bytes = 0 then
  result = 0
 elseif days = 0 then
  result = bytes
 else
  result = int(bytes / (days * 24 * 60 * 60)) 'convert days to seconds, get the ratio
 end if
 if result < 1024 then
  result = result & " bytes/sec"
 else
  result = result / 1024
  if result < 1024 then
   result = int(result) & " kilobytes/sec"
  else
   result = result / 1024
   if result < 1024 then
    result = roundto(result,2) & " megabytes/sec"
   else
    result = result / 1024
    if result < 1024 then
     result = roundto(result,2) & " gigabytes/sec" 'probably unnecessary to go past this, at least for now...
    else
    end if
   end if
  end if
 end if
 transferspeed = result
end function

function syncFolder(folderName, parent1, parent2)
 if BLACKLIST_FOLDERS.exists(folderName) then exit function
 dim result: result = 0
 dim localstarttime: localstarttime = now
 dim localbytescopied: localbytescopied = 0
 dim bytes: bytes = 0
 with createobject("scripting.filesystemobject")
  dim folder1: set folder1 = getormakefolder(pathInFolder(parent1, foldername))
  dim folder2: set folder2 = getormakefolder(pathInFolder(parent2, foldername))
 end with 
 'first, loop through ALL files in BOTH folders, to collect the full set of file names to sync
 dim file
 dim fileName
 dim fileNames: set filenames = createobject("scripting.dictionary")
 for each file in folder1.files
  filenames(file.name) = filenames(file.name) + 1
 next
 for each file in folder2.files
  filenames(file.name) = filenames(file.name) + 1
 next
 'then, sync each file name one by one
 for each filename in filenames
  bytes = syncfile(filename, folder1, folder2)
  if bytes > 0 then
   localbytescopied = localbytescopied + bytes
   result = result + bytes
  end if
 next
 'next, loop through each subfolder (in BOTH folders), to collect the full set of subfolders to sync
 dim subfolder
 dim subfoldername
 dim subfoldernames: set subfoldernames = createobject("scripting.dictionary")
 for each subfolder in folder1.subfolders
  subfoldernames(subfolder.name) = subfoldernames(subfolder.name) + 1
 next
 for each subfolder in folder2.subfolders
  subfoldernames(subfolder.name) = subfoldernames(subfolder.name) + 1
 next
 'then, call this function on each subfolder
 for each subfoldername in subfoldernames
  bytes = syncfolder(subfoldername, folder1, folder2)
  if bytes > 0 then
   localbytescopied = localbytescopied + bytes
   result = result + bytes
  end if
 next
 if result > 0 then log array("Folder """ & foldername & """ sync complete at " & transferspeed(localBytesCopied, now - localstarttime), result, timesince(localstarttime))
 syncfolder = result
end function

function syncFile(fileName, folder1, folder2)
 dim result: result = 0
 dim localStartTime: localstarttime = now
 with createobject("scripting.filesystemobject")
  dim file1: file1 = pathInFolder(folder1, filename): if .fileexists(file1) then set file1 = .getfile(file1) else set file1 = nothing
  dim file2: file2 = pathInFolder(folder2, filename): if .fileexists(file2) then set file2 = .getfile(file2) else set file2 = nothing
 end with 
 dim sourceFile
 dim destFolder 
 if file1 is nothing then
  if file2 is nothing then 'neither exist, so there is nothing to sync
   set sourcefile = nothing
   set destFolder = nothing
  else                     'only file2 exists, so use that
   set sourcefile = file2
   set destfolder = folder1
  end if
 else
  if file2 is nothing then 'only file1 eixsts, so use that
   set sourcefile = file1
   set destfolder = folder2
  else                     'both files exist, so check which is more RECENT:
   'wscript.echo file1.datelastmodified
   'wscript.echo file2.datelastmodified   
   select case true
   case file1.datelastmodified = file2.datelastmodified:      'if the files have the same modified date, both are fine and there's no reason to sync
    set sourcefile = nothing
    set destFolder = nothing
   case file1.datelastmodified > file2.datelastmodified:     'if file1 is NEWER, then we need to use THAT one
    set sourcefile = file1
    set destfolder = folder2
   case file1.datelastmodified < file2.datelastmodified: 'if file1 is OLDER, then we need to use the OTHER one                    
    set sourcefile = file2
    set destfolder = folder1
   end select
  end if
 end if  
 if file1 is nothing and file2 is nothing then
  log filename & " not found, not synced"
 elseif sourcefile is nothing or destfolder is nothing then
  'log filename & " already synced" 'this can be enabled if desired, but it mostly just clutters up the log output
 else 
  on error resume next
   sourcefile.copy slashpath(destfolder.path), true 'true will overwrite, because we've already identified which file is more important
   if err.number = 0 then 
    log array(filename & " copied to " & destfolder.path, sourcefile.size, timesince(localstarttime))
	totalfilescopied = totalfilescopied + 1
	totalBytesCopied = totalBytesCopied + sourcefile.size
	result = sourcefile.size
   else
    log array("ERROR " & err.number & ": while copying " & filename & " to " & destfolder.path, 0, timesince(localstarttime))
   end if
  on error goto 0
 end if
 syncfile = result
end function

'HELPER FUNCTIONS

function pathInFolder(folder, file)
 dim result: if isobject(folder) then result = slashpath(folder.path) else result = slashpath(folder)
 if isobject(file) then result = result & file.name else result = result & file
 pathInFolder = result
end function

function slashPath(path)
 if lastchar(path) = folder_delim then
  slashpath = path
 else
  slashpath = path & folder_delim
 end if
end function

function timeSince(beforeTime)
 dim diff: diff = now - beforetime
 dim result
 if diff >= 1 then
  result = roundto(diff, 2) & " days"
 else
  diff = diff * 24
  if diff >= 1 then
   result = roundto(diff, 2) & " hours"
  else
   diff = diff * 60
   if diff >= 1 then
    result = roundto(diff, 1) & " minutes"
   else
    diff = int(diff * 60)
	if diff = 1 then
	 result = diff & " second"
	else
	 result = diff & " seconds"
	end if
   end if  
  end if 
 end if
 timesince = result
end function

function roundTo(num, digits)
 roundto = int(num * (10 ^ digits)) / (10 ^ digits)
end function

sub getOrMakeLogSheet(logNum)
 dim result
 dim logPath: logPath = wscript.scriptfullname & ".log." & lognum & ".xlsx"
 with createobject("excel.application")
  if DEBUG_MODE then .visible = true
  if createobject("scripting.filesystemobject").fileexists(logpath) then
   on error resume next: .application.displayalerts = false
    with .workbooks.open(logpath)
	 select case err.number
     case 0: 'ignore
	 case 1004:
	  .application.quit
	  on error goto 0
	  getormakelogsheet lognum + 1
	  exit sub
	 case else:
	  .application.visible = true
	  wscript.echo err.number & ":" & err.description
	  wscript.quit
	 end select
	 set result = .worksheets(1)
	end with
   on error goto 0: .application.displayalerts = true
  else
   with .workbooks.add
    on error resume next
     .saveas logpath, 1
	 select case err.number
     case 0: 'ignore
	 case 1004:
	  wscript.echo .fullname
	  if .fullname <> logpath then
	   .application.visible = true
	   wscript.echo err.number & ":" & err.description
	   wscript.quit
	  end if
	 case else:
	  .application.visible = true
	  wscript.echo err.number & ":" & err.description
	  wscript.quit
	 end select
	on error goto 0
    set result = .worksheets(1)
   end with
  end if
  if DEBUG_MODE then result.application.visible = true
 end with
 set LOG_SHEET = result
 log "Script started": log_sheet.rows(log_sheet.usedrange.rows.count).font.bold = true
end sub

sub fatal(msg)
 log msg
 log_SHEET.application.visible = true
 log "Script cannot continue"
 wscript.echo "FATAL ERROR: " & msg
 wscript.quit
end sub

sub log(msgArrayOrString)
 dim msg
 dim msgs: if isarray(msgArrayOrString) then msgs = msgarrayorstring else msgs = array(msgarrayorstring)
 with LOG_SHEET
  dim i: i = nextRow(LOG_SHEET)
  if i > 10 then .application.activewindow.scrollrow = i - 10
  dim j: j = 1: .cells(i,j).value2 = formatdatetime(now)
  for each msg in msgs
   j = j + 1
   .cells(i,j).value2 = msg
   with .columns(j)
    if .columnwidth < MAX_LOG_COL_WIDTH then
     .autofit
	 if .columnwidth > MAX_LOG_COL_WIDTH then .columnwidth = MAX_LOG_COL_WIDTH
	end if 
   end with	
  next
  if SAVE_AFTER_EVERY_LOG then
   on error resume next 'it's possible that we are logging an error about not being able to save the file, so suppress any further error handling for this
    .parent.save
   on error goto 0
  end if
 end with
end sub

function nextRow(sheet)
 dim result
 with sheet
  if .application.worksheetfunction.counta(.usedrange.cells) = 0 then
   result = 1
  else 
   result = .usedrange.rows.count + 1
  end if
 end with
 nextrow = result
end function

function getDriveFromName(driveName)
 dim result
 for each result in createobject("scripting.filesystemobject").drives
  if result.isready then
   if ucase(result.volumename) = ucase(drivename) then exit for
  end if 
 next
 if isobject(result) then
  set getdrivefromname = result
 else
  set getdrivefromname = nothing 
 end if
end function
function subFolderPath(folder, subFolderName)
 subfolderpath = folder.path & FOLDER_DELIM & subfoldername & FOLDER_DELIM
end function
function getOrMakeFolder(folderPath)
 with createobject("scripting.filesystemobject")
  dim parentPath: parentPath = getparentfolderpath(folderpath): if not .folderexists(parentPath) then call getormakefolder(parentPath) 'will recursively make parent folders
  if not .folderexists(folderpath) then .createfolder(folderpath)
  set getormakefolder = .getfolder(folderpath)
 end with
end function
function getParentFolderPath(path)
 dim result: result = path
 if isfolderpath(result) then result = truncchars(result,1)
 result = mid(result,1,instrrev(result,folder_DELIM))
 getparentfolderpath = result
end function
function isFolderPath(path)
 isfolderpath = lastchar(path) = FOLDER_DELIM
end function
function truncChars(str, numChars)
 truncchars = mid(str,1,len(str) - numchars)
end function
function lastChar(str)
 lastChar = mid(str,len(str))
end function

'string constants

const BATOCERA  = "BATOCERA"
const BAT_SHARE = "BAT_SHARE"
