Usage: WM_position2element(elementPositioned, 
left|top|right|bottom, differenceInPixels, 
elementPositionedAgainst|window, left|top|right|bottom);
*/

  // First we set up our variables. posE is short for positionEE 
  // and posR is short for positionER.
  var posE,posR,mod;
  with(WM_position2element);
  posE = WM_checkIn(arguments[0]);
  // This block of code takes the string 'window' and makes it mean 
  // the browser window. This is handled in very different ways by 
  // Netscape and IE, so this block of code is rather long.
  if (arguments[3] == 'window') {
    if (document.all){
      posR = document.body;;
      posR.left = 0;
      posR.top = 0;
      // For some reason IE was adding four pixels. 
      // I subtracted it here.
      posR.width = document.body.offsetWidth - 4;
      posR.height = document.body.offsetHeight - 4;
    } else if (document.layers) {
      posR = document;
      posR.left = 0;
      posR.top = 0;
      posR.width = this.window.innerWidth;
      posR.height = this.window.innerHeight;
      // You need to set the width and height manually for Netscape. 
      // You can do this based on its clip.
      posE.width = posE.clip.width;
      posE.height = posE.clip.height;
    } 
  } else {
    // This is for positioning your element based on another element.
    // First, the standard checkIn procedure to conditionalize around 
    // the differences in the DOMs. You can replace this with any 
    // function that returns an object reference to a DHTML object.
    posR = WM_checkIn(arguments[3]);
    // Netscape doesn't know the object's width, only its 
    // clip.width, so I construct all that here.
    if (document.layers) {
      posE.width = posE.clip.width;
      posE.height = posE.clip.height;
      posR.width = posR.clip.width;
      posR.height = posR.clip.height;
    }
  }
  // This is where the faux properties are constructed. Right and 
  // bottom are equal to width and height, but I still use them, 
  // because it's easier to construct references to them based on 
  // the arguments later on.
  posE.right = parseInt(posE.width);
  posE.bottom = parseInt(posE.height);
  posR.right = parseInt(posR.left) + parseInt(posR.width);
  posR.bottom = parseInt(posR.top) + parseInt(posR.height);
  // This is where all that conditional work comes into play - the 
  // algorithm for the actual positioning. This is also where the 
  // difference between left and right or top and bottom is handled, 
  // through the setting of the mod[ifier] variable.
  if((arguments[1] == 'left') || (arguments[1] == 'right')) {
    if(arguments[1] == 'left') mod = 0;
    if(arguments[1] == 'right') mod = posE.right * -1;
    posE.left = parseInt(posR[arguments[4]]) + parseInt(arguments[2]) + mod;
  }
  if((arguments[1] == 'top') || (arguments[1] == 'bottom')) {
    if(arguments[1] == 'top') mod = 0;
    if(arguments[1] == 'bottom') mod = posE.bottom * -1;
    posE.top = parseInt(posR[arguments[4]]) + parseInt(arguments[2]) + mod;
  }
}