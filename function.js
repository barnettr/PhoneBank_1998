function StartForm()  {
  Criteria.SSN.focus();
}




  function slide(dir, index) {
    var elSlide = animArray[index]
    var elMenu = elSlide.parentElement
    if (!dir) {
      elMenu.style.pixelLeft-=elMenu._offset 
      if (elSlide.offsetLeft<=(-elMenu.style.pixelLeft)){
        elMenu.style.pixelLeft = -elSlide.offsetLeft
        elMenu._arrow.innerText = 4
      } else
      setTimeout("slide("+dir+","+index+")",15)
    } else
    {
      elMenu.style.pixelLeft+=elMenu._offset 
      if (elMenu.style.pixelLeft>=0){
        elMenu.style.pixelLeft = 0
        elMenu._arrow.innerText = 3
      } else
      setTimeout("slide("+dir+","+index+")",15)
    }
  }

  // Used to cache animated element
  var animArray = new Array()

  function doSlide(src) {
    el = src.parentElement
    el._offset = src.offsetLeft/10
    el._arrow = src.children.tags("SPAN")[0]
    if (el._index==null) {
      el._index = animArray.length
      animArray[animArray.length] = src
    }
    if (el.style.pixelLeft != -src.offsetLeft)
      slide(false, el._index); // Slide in
    else
      slide(true, el._index)   // Slide out
  }

