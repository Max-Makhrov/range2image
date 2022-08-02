//    ____        _   _                 
//   / __ \      | | (_)                
//  | |  | |_ __ | |_ _  ___  _ __  ___ 
//  | |  | | '_ \| __| |/ _ \| '_ \/ __|
//  | |__| | |_) | |_| | (_) | | | \__ \
//   \____/| .__/ \__|_|\___/|_| |_|___/
//         | |                          
//         |_|                          
var SaveRangeAsImageSettings = {
  folder_id: '',
  subfolder_name: 'Range2Image', // creates subfolder if set
  save2drive: true,
  measure_limit: 150, // script will assume all other rows/columns has the same size
  size_limit: 1100,   // the max. number of rows/columns,
  image_scale: 1      // will multiply image size by this number
}



//            _                 _   
//      /\   | |               | |  
//     /  \  | |__   ___  _   _| |_ 
//    / /\ \ | '_ \ / _ \| | | | __|
//   / ____ \| |_) | (_) | |_| | |_ 
//  /_/    \_\_.__/ \___/ \__,_|\__|
// makhrov.max@gmail.com
// MIT
// https://github.com/Max-Makhrov/range2image                                 
                                 



//   __  __                  
//  |  \/  |                 
//  | \  / | ___ _ __  _   _ 
//  | |\/| |/ _ \ '_ \| | | |
//  | |  | |  __/ | | | |_| |
//  |_|  |_|\___|_| |_|\__,_|
// Use this code for Google Docs, Slides, Forms, or Sheets.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('🖼️ Range2Image Menu')
      .addItem('Convert Selection to Image', 'convertSelection2Image')
      .addItem('Convert Selection to Image [SIZE x 2]', 'convertSelection2ImageX2_')
      .addItem('Convert Selection to Image [SIZE x 4]', 'convertSelection2ImageX4_')
      .addItem('Convert Selection to Image [SIZE x 8]', 'convertSelection2ImageX8_')
      .addToUi();
}
function convertSelection2ImageX2_() {
  SaveRangeAsImageSettings.image_scale = 2;
  convertSelection2Image();
}
function convertSelection2ImageX4_() {
  SaveRangeAsImageSettings.image_scale = 4;
  convertSelection2Image();
}
function convertSelection2ImageX8_() {
  SaveRangeAsImageSettings.image_scale = 8;
  convertSelection2Image();
}




//   __  __       _       
//  |  \/  |     (_)      
//  | \  / | __ _ _ _ __  
//  | |\/| |/ _` | | '_ \ 
//  | |  | | (_| | | | | |
//  |_|  |_|\__,_|_|_| |_|
function convertSelection2Image() {
  var sets = SaveRangeAsImageSettings;
  var file = SpreadsheetApp.getActive();

  /** check file is open to view */
  var fileid = file.getId();
  if (!isOpen4View_(fileid)) {
    Browser.msgBox('👀 To make the script work please share the file so that anyone can view.');
    return -1;
  }


  var range = SpreadsheetApp.getActiveRange();
  var file_name = file.getName() + '_' +  
    range.getSheet().getName() + '_' +
    range.getA1Notation();
  var url = getPdfPrintUrl_(range, sets);
  // console.log(url);

   var htmltext = HtmlService
      .createTemplateFromFile('Index')
      .evaluate()
      .getContent();
  
  // add the function names 
  htmltext = htmltext.replace(/IMPORT_PDF_URL/m, url);
  var scale = sets.image_scale;
  htmltext = htmltext.replace(/IMAGE_SCALE/m, scale);
  htmltext = htmltext.replace(/IMAGE_NAME/m, file_name);

  var html = HtmlService.createTemplate(htmltext).evaluate()
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

  SpreadsheetApp.getUi()
      .showModalDialog(html, '🖼️Range2Image');

}




//   _____    _  __ 
//  |  __ \  | |/ _|
//  | |__) |_| | |_ 
//  |  ___/ _` |  _|
//  | |  | (_| | |  
//  |_|   \__,_|_|  
/**
 * converts range to Url
 * ready to be saved as PDF
 * 
 * More info here:
 * https://stackoverflow.com/questions/46088042
 * https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
 * https://kandiral.ru/googlescript/eksport_tablic_google_sheets_v_pdf_fajl.html
 */
function getPdfPrintUrl_(range, options) {
  var ratio = 96; // get inch from pixel

  range = range || SpreadsheetApp.getActiveRange();
  var sheet = range.getSheet();
  var file = SpreadsheetApp.getActive();

  var fileurl = file.getUrl();
  var sheetid = sheet.getSheetId();
  var rownum = range.getRow();
  var columnnum = range.getColumn();
  var rownum2 = range.getLastRow();
  var columnnum2 = range.getLastColumn();

  if ((rownum2-rownum+1) > options.size_limit) {
    throw '😢The range exceeded the limit of ' +  options.size_limit + ' rows';
  }
  if ((columnnum2-columnnum+1) > options.size_limit) {
    throw '😢The range exceeded the limit of ' +  options.size_limit + ' columns';
  }

  file.toast('Please wait...', '📐Measuring Range...');
  // get width in pixels 
  var w = 0, size;
  for (var i = columnnum; i <= columnnum2; i++) {
    if (i <= options.measure_limit) {
      size = sheet.getColumnWidth(i);
    }
    w += size;
    if ((i % 50) === 0 && i <= options.measure_limit) {
      file.toast(
        'Done ' + i + ' columns of ' + columnnum2,
        '↔📐Measuring width...');
    }
  }
  if (i > options.measure_limit) {
    file.toast(
      'Estimation: all other columns are the same size',
      '↔📐Measuring width...');
  }
 
  // get row height in pixels
  var h = 0;
  for (var i = rownum; i <= rownum2; i++) {
    if (i <= options.measure_limit) {
      size = sheet.getRowHeight(i);
    }
    h += size
    /** manual correction */
    if (size === 2) {
      h-=1;
    } else {
      // h -= 0.42; /** TODO → test the range to make it fit any range */
    }
    
    if ((i % 50) === 0 &&  i <= options.measure_limit) {
      file.toast(
        'Done ' + i + ' rows of ' + rownum2,
        '↕📐Measuring height...');
    }
  }
  if (i > options.measure_limit) {
    file.toast(
      'Estimation: all other rows are the same size',
      '↕📐Measuring height...');
  }

  var hh = Math.round(h/ratio * 1000) / 1000;
  var ww = Math.round(w/ratio * 1000) / 1000;

  // Browser.msgBox(
  //   JSON.stringify(
  //     {
  //       h: h,
  //       w: w,
  //       hh: hh,
  //       ww: ww
  //     }, null, 2
  //   )
  // );
  
  var sets = {
    url:      fileurl,
    sheetId:  sheetid,
    r1:       rownum-1,
    r2:       rownum2,
    c1:       columnnum-1,
    c2:       columnnum2,
    size:     ww +'x' + hh,          //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
    // portrait: true,       //true= Potrait / false= Landscape
    scale: 2,          //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
    top_margin: 0,     //All four margins must be set!        
    bottom_margin: 0,  //All four margins must be set!       
    left_margin: 0,    //All four margins must be set!         
    right_margin: 0,   //All four margins must be set!
  }
  var rangeParam =
      '&r1=' + sets.r1 + 
      '&r2=' + sets.r2 + 
      '&c1=' + sets.c1 +
      '&c2=' + sets.c2;
  var sheetParam = '&gid=' + sets.sheetId;
  var isPortrait = '';
  if (sets.portrait) {
    //true= Potrait / false= Landscape
    isPortrait = '&portrait='      + sets.portrait;
  }
  var exportUrl = sets.url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size='          + sets.size             //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
      +  isPortrait     
      + '&scale='         + sets.scale            //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page     
      + '&top_margin='    + sets.top_margin       //All four margins must be set!       
      + '&bottom_margin=' + sets.bottom_margin    //All four margins must be set!     
      + '&left_margin='   + sets.left_margin      //All four margins must be set! 
      + '&right_margin='  + sets.right_margin     //All four margins must be set!     
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + 'horizontal_alignment=LEFT' // //LEFT/CENTER/RIGHT
      + '&gridlines=false'
      + "&fmcmd=12"
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam;
  // console.log('exportUrl=' + exportUrl);
  // Browser.msgBox(exportUrl);

  return exportUrl;
}




//      /\                         
//     /  \   ___ ___ ___  ___ ___ 
//    / /\ \ / __/ __/ _ \/ __/ __|
//   / ____ \ (_| (_|  __/\__ \__ \
//  /_/    \_\___\___\___||___/___/
// function test_isOpen4View() {
//   var ss = SpreadsheetApp.getActive();
//   var fileid = ss.getId();
//   console.log(isOpen4View_(fileid));
// }
/**
 * test if anyone can view the file
 * 
 * credits:
 * https://www.labnol.org/code/19538-file-sharing-permissions
 */
function isOpen4View_(fileid) {

  var file = DriveApp.getFileById(fileid);
  var access = file.getSharingAccess();
  var privacy;

  switch (access) {
    case DriveApp.Access.PRIVATE:
      privacy = 'Private';
      break;
    case DriveApp.Access.ANYONE:
      privacy = 'Anyone';
      break;
    case DriveApp.Access.ANYONE_WITH_LINK:
      privacy = 'Anyone with a link';
      break;
    case DriveApp.Access.DOMAIN:
      privacy = 'Anyone inside domain';
      break;
    case DriveApp.Access.DOMAIN_WITH_LINK:
      privacy = 'Anyone inside domain who has the link';
      break;
    default:
      privacy = 'Unknown';
  }

  // console.log(privacy);

  var result = [
    'Anyone', 'Anyone with a link'
  ].indexOf(privacy) > -1;

  return result;
}




//   _____                            ___  _____       _           
//  |_   _|                          |__ \|  __ \     (_)          
//    | |  _ __ ___   __ _  __ _  ___   ) | |  | |_ __ ___   _____ 
//    | | | '_ ` _ \ / _` |/ _` |/ _ \ / /| |  | | '__| \ \ / / _ \
//   _| |_| | | | | | (_| | (_| |  __// /_| |__| | |  | |\ V /  __/
//  |_____|_| |_| |_|\__,_|\__, |\___|____|_____/|_|  |_| \_/ \___|
//                          __/ |                                  
//                         |___/                                   
/**
 * Callback function from Index.html
 */
function saveDataUrlToFolder(dataURI, image_name) {
  var sets = SaveRangeAsImageSettings;
  if (!sets.save2drive) {
    return '✔️Done!<BR>Rigth clisck the image to save locally ↓';
  }
  sets.file_name = image_name;
  try {
    return saveDataUrlToFolder_(dataURI, sets);
  } catch (e) {
    return e.message + '; stack: ' + e.stack;
  }

}
function saveDataUrlToFolder_(dataURI, sets) {
  var root;
  if (sets.folder_id === '') {
    root = DriveApp.getRootFolder();
  } else {
      try {
        root = DriveApp.getFolderById(sets.folder_id);
      } catch (err) {
        return err;
      }
    }
    if (!root) {
      return 'no folder with id = ' + sets.folder_id;
    }
    var folder;
    if (sets.subfolder_name === '' || !sets.subfolder_name) {
      folder = root;
    } else {
      folder = createFolderInFolder_(root, sets.subfolder_name);
    }
    

    var type = (dataURI.split(";")[0]).replace('data:','');
    var imageUpload = Utilities.base64Decode(dataURI.split(",")[1]);
    var blob = Utilities.newBlob(imageUpload, type, sets.file_name + ".png");

    try {
      var file = folder.createFile(blob);
    } catch (err) {
      return 'Oops! Range is too big and cannot be rendered';
    }

    var result = '✔️Image is saved to Drive!<BR>' + 
      '<a href="' + file.getUrl() + '">' + sets.file_name + '</a>'  +
      '<BR><BR> Rigth click the image to save locally ↓'

    return result;
}




//   ______    _     _               
//  |  ____|  | |   | |              
//  | |__ ___ | | __| | ___ _ __ ___ 
//  |  __/ _ \| |/ _` |/ _ \ '__/ __|
//  | | | (_) | | (_| |  __/ |  \__ \
//  |_|  \___/|_|\__,_|\___|_|  |___/
/**
 * create folder in folder
 * 
 * @param {DriveApp.Folder} folder
 * @param {string} name
 */
function createFolderInFolder_(folder, name) {
  var existingFolder = isFolderInFolder_(folder, name);
  // exclude creating folders with the same names
  if (existingFolder) { return existingFolder; } 
  var result = folder.createFolder(name);
  return result;  
}
/**
 * folder with name exists
 * 
 * @param {DriveApp.Folder} folder
 * @param {string} name
 */
function isFolderInFolder_(folder, name) {
  var folders = folder.getFolders();
  var folder;
  while (folders.hasNext()) {
    folder = folders.next();
    if (folder.getName() === name) { return folder; }
  }
  return false; 
}
