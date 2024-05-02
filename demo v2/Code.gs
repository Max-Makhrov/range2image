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
}



//            _                 _   
//      /\   | |               | |  
//     /  \  | |__   ___  _   _| |_ 
//    / /\ \ | '_ \ / _ \| | | | __|
//   / ____ \| |_) | (_) | |_| | |_ 
//  /_/    \_\_.__/ \___/ \__,_|\__|
// makhrov.max@gmail.com
// https://twitter.com/max__makhrov
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
      .addItem("Convert selection to image", "convertSelection2Image")
      .addToUi();
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

  var range = SpreadsheetApp.getActiveRange();
  var file_name = file.getName() + '_' +  
    range.getSheet().getName() + '_' +
    range.getA1Notation();
  var url = getPdfPrintUrl_(range, sets);

  var blob = UrlFetchApp.fetch(url, {headers: {authorization: "Bearer " + ScriptApp.getOAuthToken()}}).getBlob();
  var imageBlob = convertPDFToPNG_(blob);

  // As a sample, create PNG images as PNG files.
  var folderId = getFolderByNameCreateIfNotExists_(sets.subfolder_name, sets.folder_id);
  createPngInFolder_(imageBlob, file_name, folderId);
}



//   _____                        ___  _____    _  __ 
//  |  __ \                      |__ \|  __ \  | |/ _|
//  | |__) |__ _ _ __   __ _  ___   ) | |__) |_| | |_ 
//  |  _  // _` | '_ \ / _` |/ _ \ / /|  ___/ _` |  _|
//  | | \ \ (_| | | | | (_| |  __// /_| |  | (_| | |  
//  |_|  \_\__,_|_| |_|\__, |\___|____|_|   \__,_|_|  
//                      __/ |                         
//                     |___/                          
/**
 * converts range to Url
 * ready to be saved as PDF
 * 
 * More info here:a
 * https://stackoverflow.com/questions/46088042
 * https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
 * https://kandiral.ru/googlescript/eksport_tablic_google_sheets_v_pdf_fajl.html
 * 
 * @param {int} options.size_limit    1100 rows/columns (in tests)
 * @param {int} options.measure_limit 150 rows/columns
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

  // add 0.1 inch to fit some ranges
  var hh = Math.round(h/ratio * 1000 + 100) / 1000;
  var ww = Math.round(w/ratio * 1000 + 100) / 1000;
  
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

  return exportUrl;
}





//   _____    _  __ ___  _____             
//  |  __ \  | |/ _|__ \|  __ \            
//  | |__) |_| | |_   ) | |__) | __   __ _ 
//  |  ___/ _` |  _| / /|  ___/ '_ \ / _` |
//  | |  | (_| | |  / /_| |   | | | | (_| |
//  |_|   \__,_|_| |____|_|   |_| |_|\__, |
//                                    __/ |
//                                   |___/ 
/**
 * https://stackoverflow.com/a/55152707/5372400
 * https://stackoverflow.com/questions/75315605
 * 
 * @param {Blob} blob
 * 
 * @returns {Blob}
 */
function convertPDFToPNG_(blob) {
  var fileId = createFileInFolder_("TMP_DELETE.pdf", blob, "application/pdf");

  var link = null;
  var maxTries = 5;
  var tries = 1;
  while (!link) {
    if (tries > maxTries) {
      throw new Error("Could not get link in " + maxTries + " seconds");
    }
    Utilities.sleep(1000);
    link = Drive.Files.get(fileId, { fields: "thumbnailLink" }).thumbnailLink;
    tries++;
  }

  var thumbnailURL = link.replace(/=s.+/, "=s2500");
  console.log(thumbnailURL);
  var pngBlob = UrlFetchApp.fetch(thumbnailURL).getBlob();
  Drive.Files.remove(fileId);
  return pngBlob;
}




//   ______    _     _           ____        _   _                      
//  |  ____|  | |   | |         |  _ \      | \ | |                     
//  | |__ ___ | | __| | ___ _ __| |_) |_   _|  \| | __ _ _ __ ___   ___ 
//  |  __/ _ \| |/ _` |/ _ \ '__|  _ <| | | | . ` |/ _` | '_ ` _ \ / _ \
//  | | | (_) | | (_| |  __/ |  | |_) | |_| | |\  | (_| | | | | | |  __/
//  |_|  \___/|_|\__,_|\___|_|  |____/ \__, |_| \_|\__,_|_| |_| |_|\___|
//                                      __/ |                           
//                                     |___/                            
/**
 * @prop {String} folderName
 * @prop {String} [parentFolderId]
 * 
 * @returns {String} folderId
 */
function getFolderByNameCreateIfNotExists_(folderName, parentFolderId) {
  if (parentFolderId === "") parentFolderId = null;
  var mimeTypeStr = 'application/vnd.google-apps.folder';
  var q = "name = '" + folderName + "' and mimeType = '" + mimeTypeStr + "'";
  var parentId = parentFolderId || "root";
  q += " and ('" + parentId + "' in parents)";
  var searchFolders = Drive.Files.list({q: q})
  if (searchFolders.files) {
    if (searchFolders.files.length) {
      return searchFolders.files[0].id;
    }
  }
  var newFolder = {
    name: folderName,
    mimeType: mimeTypeStr
  };
  if (parentFolderId) {
    newFolder.parents = [parentFolderId];
  }
  var folder = Drive.Files.create(newFolder);
  return folder.id;
}



//   _____          _______    ______    _     _           
//  |  __ \        |__   __|  |  ____|  | |   | |          
//  | |__) | __   __ _| | ___ | |__ ___ | | __| | ___ _ __ 
//  |  ___/ '_ \ / _` | |/ _ \|  __/ _ \| |/ _` |/ _ \ '__|
//  | |   | | | | (_| | | (_) | | | (_) | | (_| |  __/ |   
//  |_|   |_| |_|\__, |_|\___/|_|  \___/|_|\__,_|\___|_|   
//                __/ |                                    
//               |___/          
/**
 * @param {Blob} blob
 * @param {String} fileName
 * @param {String} folderId
 */
function createPngInFolder_(blob, fileName, folderId) {
  mimeType = "image/x-png";
  createFileInFolder_(fileName, blob, mimeType, folderId);
}


/**
 * https://developers.google.com/drive/api/reference/rest/v3/files#File
 * @param {String} fileName
 * @param {Blob} blob
 * @param {String} mimeType
 * @param {String} [folderId]
 * 
 * @returns {String} id
 */
function createFileInFolder_(fileName, blob, mimeType, folderId) {
  var file = {
    mimeType: mimeType,
    name: fileName
  };
  if (folderId && folderId !== "") {
    file.parents = [ folderId ];
  }
  var file = Drive.Files.create(file, blob);
  return file.id;
}
