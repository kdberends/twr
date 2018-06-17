// Move file to another Google Drive Folder
function moveFileToAnotherFolder(fileID, targetFolderID) {

  var file = DriveApp.getFileById(fileID);
  
  // Remove the file from all parent folders
  var parents = file.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    parent.removeFile(file);
  }

  DriveApp.getFolderById(targetFolderID).addFile(file);
  
}

function getParentFolderOfFile(file){
 return DriveApp.getFileById(file.getId()).getParents().next().getId()
}


function getSubfolder(foldername, parentfolder){
  folders = DriveApp.getFolderById(parentfolder).getFolders()
  while (folders.hasNext()) {
   var folder = folders.next();
    if (folder.getName() == foldername){
      return folder.getId()
    }

  }
  return null
}