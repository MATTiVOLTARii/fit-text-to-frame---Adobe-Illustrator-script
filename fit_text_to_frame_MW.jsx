// fit_text_to_frame.jsx

// original code by hhas01 https://community.adobe.com/t5/illustrator-discussions/auto-change-text-size-to-size-of-text-box/m-p/12535330#M299271
// modified by matthias wrembel

// resizes text larger/smaller to fit the selected text frame(s)
// (the text frame must be fully selected; point text and non-text items are ignored)
// current limitations: this assumes all characters are same size; leading is not adjusted

(function() {

var MIN_SIZE = 1, MAX_SIZE = 1296 // practical limits, in pts
var MIN_STEP = 0.05 // the smaller this is, the more precise (and slower) the fit

function fitTextToFrame(textFrame) {
    // textFrame : TextFrame -- path or area text
    if (textFrame.typename !== 'TextFrame' || textFrame.kind === TextType.POINTTEXT) { return }

    var l=0
    var countlines=textFrame.lines.length
    while(l< countlines){

        var length = textFrame.lines[l].contents.replace(/\s+$/, "").length // length of printable text (ignores trailing whitespace)
        if (length === 0) { return } // skip empty text frame

    var visLength = textFrame.lines[l].contents.replace(/\s+$/, "").length // length of written text
    var style = textFrame.textRange.lines[l].characterAttributes
    var size = style.size



    // 1. resize text so it is just overflowing
    while ((visLength > length) && size >= MIN_SIZE * 2) { // decrease size until no overflow
        size /= 2.0
        style.size = size
        visLength = textFrame.lines[l].contents.replace(/\s+$/, "").length
        
    }
   
   if ((visLength > length)) { return } // hit minimum size and it's still overflowing, so give up
    while ((visLength < length) && size <= MAX_SIZE / 2) { // increase size until text overflows
        size *= 2.0
        style.size = size
        visLength = textFrame.lines[l].contents.replace(/\s+$/, "").length
      
    }
    
   if ((visLength < length)) { return } // hit maximum size and it hasn't overflowed, so give up

    
    // 2. now do binary search for the largest size that will just fit the text frame
    var step = size / 1.5
    while (step > MIN_STEP) {
        step /= 1.25        
        size += (visLength < length) ? -step : step
        style.size = size
        visLength = textFrame.lines[l].contents.replace(/\s+$/, "").length
    
    }
   
    // 3. if text ends up overflowing after last step, step it back down
    while ((visLength < length)) { 
        style.size = size - step
        step *= 1.1
        visLength = textFrame.lines[l].contents.replace(/\s+$/, "").length

    }
    
    
    l++ 
}
}

var items = app.activeDocument.selection
for (var i = 0; i < items.length; i++) { // iterate over currently selected items
    fitTextToFrame(items[i])
}

})()