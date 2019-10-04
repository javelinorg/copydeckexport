/*
 *
 * Copy_Deck_Export 1.6 - updated text extract script by R!chard Silva http://richardsilva.info
 * parts of this script orignated from Bramus! - http://www.bram.us/
 * text layer style export added by Max Glenister (@omgmog) http://www.omgmog.net
 *
 * v 1.6 - 2019.09.19 - moved file creation down inside function to create one file per top layersets
 * v 1.5 - 2019.07.08 - added Key words to force extract
 * v 1.4 - 2019.07.03 - now works with Photoshop CC 2018
 *                    - added config options
 * v 1.3 - 2008.03.16 - Base rewrite, now gets all layers (sets & regular ones) in one variable.
 *                    - Layer Path & Layer Name in export
 *                    - Cycle Multiple Files
 * v 1.2 - 2008.03.11 - User friendly version Added filesave dialog (*FIRST RELEASE*)
 * v 1.1 - 2008.03.11 - Extended version,   Loops sets too (I can haz recursiveness)
 * v 1.0 - 2008.03.11 - Basic version,    Loops all layers (no sets though)
 *
 * Licensed under the Creative Commons Attribution 2.5 License - http://creativecommons.org/licenses/by/2.5/
 *
 */
/**
 * CONFIG - CHANGE THESE IF YOU LIKE
 * -------------------------------------------------------------
 */
// Use save as dialog (true/false)?
// This will be overriden to false when running TextExport on multiple files!
var useDialog = true;

// Open file when done (true/false)?
// This will be overriden to true when running TextExport on multiple files!
var openFile  = false; // currently commented out (around line 166) due to dumping out lots of files per 'art board'

// Create one extract per top layer?
var filePerArtboard = true;

// Sort the LayerSets based on path
var sortLayerSets = false;

// show only visable layers
var onlyVisableLayers = true;

// dump all layers (not just text)
var allLayerTypes = false;

// dump all layer attributes (beyond the predefined)
var allLayerAttributes = false;

// text separator
var separator = "*************************************";

/**
 * NO NEED TO CHANGE ANYTHING BENEATH THIS LINE
 * -------------------------------------------------------------
 */
/**
 *  TextExport Init function
 * -------------------------------------------------------------
 */
var scriptTitle = 'CopyDeckExport';
var filePath = '';
var fileOut;
// this array contains layer names that get extracted
// regardless of visability
var copyDeckKeys = ['subjectline',
                    'preheader',
                    'previewtext',
                    'greeting',
                    'alttext',
                    'jobinfo',
                    ]

// Progress Window Stuffs
var progressValue = 10;
var progressWindow = new Window("palette{text:'Please be patient...',bounds:[100,100,580,140],progress:Progressbar{bounds:[20,10,460,30] , minvalue:0,value:" + progressValue + "}};" );
var progressD = progressWindow.graphics;
progressD.backgroundColor = progressD.newBrush(progressD.BrushType.SOLID_COLOR, [0.00, 0.00, 0.00, 1]);

// start export by setting up a few things
function initTextExport(){
  if($.os.search(/windows/i) != -1){
    fileLineFeed = "windows";
  }else{
    fileLineFeed = "macintosh";
  }

  // Do we have a document open?
  if(app.documents.length < 1){
    alert("Please open at least one file", scriptTitle + ' Error', true);
    return;
  }

  // Oh, we have more than one document open!
  if(app.documents.length > 1){
    var runMultiple = confirm(scriptTitle + " has detected Multiple Files.\nDo you wish to run " + scriptTitle + " on all opened files?", true, scriptTitle);
    if(runMultiple === true){
      docs = app.documents;
    }else{
      docs = [app.activeDocument];
    }
  }else{
    runMultiple   = false;
    docs    = [app.activeDocument];
  }
  // Loop all documents
  for(var i = 0; i < docs.length; i++){
    // the default output file name
    var outputFileName = docs[i].name;
    // useDialog (but not when running multiple
    if((runMultiple !== true) && (useDialog === true)){
      // Pop up save dialog
      //var saveFile = File.saveDialog("Please select a file to export the text extract to:", outputFileName);
      var saveFile = File.saveDialog('Please selecte an output directory.' + "\n" + 'Disregard the file name. It will get overwriten.');

      // User Cancelled
      if(saveFile == null){
        alert('Don\'t want to extract text?!' + "\n" + 'No worries. You can do this again later.');
        return;
      }
      // set filePath and fileName to the one chosen in the dialog
      //filePath = saveFile.path + "/" + saveFile.name;
      filePath = saveFile.path + '/' + outputFileName;
    }else{
      // Don't use Dialog
      // Auto set filePath and fileName
      filePath = Folder.myDocuments + '/' + outputFileName ;
    }

    // Progress Window Stuffs
    progressWindow.center();
    progressWindow.show();

    // // create outfile
    // fileOut = new File(filePath);

    // // clear dummyFile
    // dummyFile = null;

    // // set linefeed
    // fileOut.linefeed = fileLineFeed;

    // // open for write
    // fileOut.open("w", "TEXT", "????");
    // fileOut.encoding = "UTF-8";

    // // Append title of document to file
    // fileOut.writeln(separator);
    // fileOut.writeln('* START ' + scriptTitle + ' for ' + docs[i].name);

    // Set active document
    app.activeDocument = docs[i];

    // call to the core with the current document
    //goTextExport2(app.activeDocument, fileOut, '/');
    copyDeckExport(app.activeDocument, null, '/',0);

    // //  Hammertime!
    // fileOut.writeln(separator);
    // fileOut.writeln('* FINISHED ' + scriptTitle + ' for ' + docs[i].name);
    // fileOut.writeln(separator);

    // // close the file
    // fileOut.close();

    // Give notice that we're done or open the file (only when running 1 file!)
    if(runMultiple === false){
      if(openFile === true){
        //fileOut.execute();
      }else{
        alert("File was saved to:\n" + Folder.decode(filePath), scriptTitle);
      }
    }
  }

  if(runMultiple === true){
    alert('Parsed ' + documents.length + " files;\nFiles were saved in your documents folder", scriptTitle);
  }
}

/**
 * Sort the layers array based on name
 * -------------------------------------------------------------
 */
function sortLayers(layers){
  //var layers = activeDocument.layers;
  var layersArray = [];
  var len = layers.length;
  // store all layers in an array
  for(var i = 0; i < len; i++){
    layersArray.push(layers[i]);
  }

  // sort layer top to bottom
  // based on the layer path
  layersArray.sort();

  // This for loop will move the layers within the PSD
  // for(i = 0; i < len; i++){
  //   layersArray[i].move(layers[i], ElementPlacement.PLACEBEFORE);
  // }
  return layersArray;
}

/**
 * This will search the given layer name and look to see if
 * it's one of the key copy deck layers that needs to be extracted
 */
function searchCopyKeys(layerName){
  // strip the layer name to reduce the chance of
  // someone using mix case or spaces or whatever
  var clearnLayerName = layerName.replace(/[^a-z0-9]+/gi, '').toLowerCase();
  for(var x=0; x<copyDeckKeys.length; x++){
    if(clearnLayerName.indexOf(copyDeckKeys[x]) > -1){
      return true;
    }
  }
}
/**
 * CopyDeckExport Core Function (V1)
 * -------------------------------------------------------------
 */
function copyDeckExport(el, fileOut, path, depth){
  // Get the layers
  var layers = el.layers;

  // sort the layers
  // but only if you want to
  if(sortLayerSets){
    layers = sortLayers(layers);
  }


  // set the max progres bar
  // progressWindow.progress.maxvalue = layers.length;

  // Loop the layers
  //for(var layerIndex = layers.length; layerIndex > 0; layerIndex--){
  for(var layerIndex = 0; layerIndex < layers.length; layerIndex++){

    // Progress Window Stuffs
    progressWindow.progress.value++;
    progressWindow.layout.layout(true);

    // curentLayer ref
    var currentLayer = layers[layerIndex];
    // lets see if current layer is visable

    // currentLayer is a LayerSet
    if((currentLayer.typename == "LayerSet") && (currentLayer.visible)){
      if(depth == 0){
        // create outfile
        fileOut = new File(filePath + '_' + currentLayer.name + '.txt');

        // set linefeed
        fileOut.linefeed = fileLineFeed;

        // open for write
        fileOut.open("w", "TEXT", "????");
        fileOut.encoding = "UTF-8";

        // Append title of document to file
        fileOut.writeln(separator);
        fileOut.writeln('* START ' + scriptTitle + ' for ' + currentLayer.name);
      }
      copyDeckExport(currentLayer, fileOut, path + currentLayer.name + '/', depth+1);
    // currentLayer is not a LayerSet
    }else{
      var layerName = currentLayer.name;

      if((currentLayer.visible) || (searchCopyKeys(layerName))){
        if((allLayerTypes) && (currentLayer.kind != LayerKind.TEXT)){
          fileOut.writeln('');
          fileOut.writeln(separator);
          fileOut.writeln('');
          fileOut.writeln('LayerPath: ' + path);
          fileOut.writeln('LayerName: ' + currentLayer.name);
          fileOut.writeln('');
          for(var key in currentLayer){
            try{
              if(allLayerAttributes){
                fileOut.writeln('* ' + key + ': ' + currentLayer[key]);
              }
            }catch(err){
              if(allLayerAttributes){
                fileOut.writeln('* ' + key + ': undefined');
              }
            }
          }
        }
        // Layer is visible and Text --> we can haz copy paste!
        if(currentLayer.kind == LayerKind.TEXT){
          fileOut.writeln('');
          fileOut.writeln(separator);
          fileOut.writeln('');
          fileOut.writeln('LayerPath: ' + path);
          fileOut.writeln('LayerName: ' + currentLayer.name);
          fileOut.writeln('');
          fileOut.writeln('LayerContent:');
          fileOut.writeln('');
          try{
            fileOut.writeln(currentLayer.textItem.contents);
          }catch(err){
            fileOut.writeln('ERROR pulling content!');
            fileOut.writeln(err);
            fileOut.writeln('END - pulling content!');
          }

          fileOut.writeln('');
          // additional exports added by Max Glenister for font styles
          if(currentLayer.textItem.contents){
            fileOut.writeln('LayerStyles:');
            fileOut.writeln('');
            var textItemProps = currentLayer.textItem;
            for(var key in textItemProps){
              try{
                if(allLayerAttributes){
                  fileOut.writeln('* ' + key + ': ' + textItemProps[key]);
                }
//                 switch(key){
//                   case 'color':
// //                    fileOut.writeln('* ' + key + ': #' + (currentLayer.textItem.color.rgb.hexValue?currentLayer.textItem.color.rgb.hexValue:''));
//                     break;
//                   case 'capitalization':
// //                    fileOut.writeln('* ' + key + ': ' + (textItemProps[key]=="TextCase.NORMAL"?"normal":"uppercase"));
//                     break;
// //                  case 'direction':
// //                    fileOut.writeln('* ' + key + ': ' + (textItemProps[key]=="Direction.HORIZONTAL"?"horizontal":"vertical"));
// //                    break;
//                   case 'font':
// //                  case 'leading':
//                   case 'size':
// //                  case 'position':
//                   case 'justification':
// //                    fileOut.writeln('* ' + key + ': ' + textItemProps[key]);
//                     break;
// //                  case 'fauxBold':
// //                  case 'fauxItalic':
// //                    fileOut.writeln('* ' + key + ': ' + (textItemProps[key]?'true':'fasle'));
// //                    break;
//                   default:
//                     // if(allLayerAttributes){
//                     //   fileOut.writeln('* ' + key + ': ' + textItemProps[key]);
//                     // }
//                 }
              }catch(err){
                if(allLayerAttributes){
                  fileOut.writeln('* ' + key + ': undefined');
                }
              }
            }
          }
        }
      }else{
        // layer is NOT visible and will be skipped
        // fileOut.writeln('');
        // fileOut.writeln(separator);
        // fileOut.writeln('');
        // fileOut.writeln('LayerPath: ' + path);
        // fileOut.writeln('LayerName: ' + currentLayer.name);
        // fileOut.writeln('');
        // fileOut.writeln('Layer NOT visible and therefor not exported.');
        // fileOut.writeln('');
      }
    }
  }
}

/**
 * TextExport Core Function (V2)
 * -------------------------------------------------------------
 */
function goTextExport2(el, fileOut, path){
  // Get the layers
  var layers = el.layers;
  // Loop the layers
  for(var layerIndex = layers.length; layerIndex > 0; layerIndex--){

    // curentLayer ref
    var currentLayer = layers[layerIndex-1];

    // currentLayer is a LayerSet
    if(currentLayer.typename == "LayerSet"){
      goTextExport2(currentLayer, fileOut, path + currentLayer.name + '/');
    // currentLayer is not a LayerSet
    }else{
      // Layer is visible and Text --> we can haz copy paste!
      if((currentLayer.visible) && (currentLayer.kind == LayerKind.TEXT)){
        fileOut.writeln(separator);
        fileOut.writeln('');
        fileOut.writeln('LayerPath: ' + path);
        fileOut.writeln('LayerName: ' + currentLayer.name);
        fileOut.writeln('');
        fileOut.writeln('LayerContent:');
        fileOut.writeln(currentLayer.textItem.contents);
        fileOut.writeln('');
        // additional exports added by Max Glenister for font styles
        if(currentLayer.textItem.contents){
          fileOut.writeln('LayerStyles:');
//              fileOut.writeln('* capitalization: '+(currentLayer.textItem.capitalization=="TextCase.NORMAL"?"normal":"uppercase"));
          fileOut.writeln('* color: #'+(currentLayer.textItem.color.rgb.hexValue?currentLayer.textItem.color.rgb.hexValue:''));
//              fileOut.writeln('* fauxBold: '+(currentLayer.textItem.fauxBold?currentLayer.textItem.fauxBold:''));
//              fileOut.writeln('* fauxItalic: '+(currentLayer.textItem.fauxItalic?currentLayer.textItem.fauxItalic:''));
          fileOut.writeln('* font: '+currentLayer.textItem.font);
          //fileOut.writeln('leading: '+(currentLayer.textItem.leading=='auto-leading'?'auto':currentLayer.textItem.leading));
          fileOut.writeln('* size: '+currentLayer.textItem.size);
//              fileOut.writeln('* tracking: '+(currentLayer.textItem.fauxItalic?currentLayer.textItem.fauxItalic:''));
          fileOut.writeln('');
        }
      }
    }
  }
}
/**
 *  TextExport Boot her up
 * -------------------------------------------------------------
 */
initTextExport();