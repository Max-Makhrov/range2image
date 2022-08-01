var SaveRangeAsImageSettings = {
  folder_id: '',
  save2drive: true,
  measure_limit: 150, // script will assume all other rows/columns has the same size
  size_limit: 1100,   // the max. number of rows/columns,
  image_scale: 1      // will multiply image size by this number
}


// Use this code for Google Docs, Slides, Forms, or Sheets.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('🖼️ Range2Pic Menu')
      .addItem('Convert Selection to Image', 'convertSelection2Image')
      .addItem('Convert Selection to Image [SIZE x 2]', 'convertSelection2ImageX2_')
      .addItem('Convert Selection to Image [SIZE x 2]', 'convertSelection2ImageX4_')
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



function convertSelection2Image() {
  var file = SpreadsheetApp.getActive();
  var range = SpreadsheetApp.getActiveRange();
  var file_name = file.getName() + '_' + range.getA1Notation();
  SaveRangeAsImageSettings.file_name = file_name;
  var url = getPdfPrintUrl_(range, SaveRangeAsImageSettings);
  // console.log(url);


   var htmltext = HtmlService
      .createTemplateFromFile('Index')
      .evaluate()
      .getContent();
  
  // add the function names 
  htmltext = htmltext.replace(/IMPORT_PDF_URL/m, url);
  var scale = SaveRangeAsImageSettings.image_scale;
  htmltext = htmltext.replace(/IMAGE_SCALE/m, scale);

  var html = HtmlService.createTemplate(htmltext).evaluate()
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Range2Pic');

}


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
 
  // get row height in pixelsh
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
      // h -= 0.415;
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
      // + '&fzr=FALSE'      
      + sheetParam
      + rangeParam;
  // console.log('exportUrl=' + exportUrl);
  // Browser.msgBox(exportUrl);

  return exportUrl;
}




/**
 * Callback function from Index.html
 */
function saveDataUrlToFolder(dataURI) {
  var sets = SaveRangeAsImageSettings;
  
  if (!sets.save2drive) {
    return 'Done!';
  }

  try {
    return saveDataUrlToFolder_(dataURI, sets);
  } catch (e) {
    return e.message + '; stack: ' + e.stack;
  }

}
function saveDataUrlToFolder_(dataURI, sets) {
  var folder;
  if (sets.folder_id === '') {
    folder = DriveApp.getRootFolder();
  } else {
      try {
        folder = DriveApp.getFolderById(sets.folder_id);
      } catch (err) {
        return err;
      }
    }
    if (!folder) {
      return 'no folder with id = ' + sets.folder_id;
    }

    var type = (dataURI.split(";")[0]).replace('data:','');
    var imageUpload = Utilities.base64Decode(dataURI.split(",")[1]);
    var blob = Utilities.newBlob(imageUpload, type, "nameOfImage.png");

    try {
      var file = folder.createFile(blob);
    } catch (err) {
      return 'Oops! Range is too big and cannot be rendered';
    }
    

    return 'Image is saved to Drive! Your URL:<BR>' + file.getUrl();
}
