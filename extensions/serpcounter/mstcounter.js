function mstSerpMain() {
    var elements = document.querySelectorAll('#search .yuRUbf');
   //var results = [...elements].filter(element => !element.closest('.LC20lb'));
   // var _g = document.querySelector('.g');
   var results = document.getElementsByClassName('yuRUbf'); /* main */

        // 2+ matches = first match is page number
        var currentPage = 1;
       
        var resultsPerPage = results.length;

    // Reset localStorage on page 1 for new searches
    if (currentPage == 1) {
        localStorage.removeItem('countLastPage');
        localStorage.removeItem('lastPage');
    }

    var countDisplay = 1;
    for (var countActual = 0; countActual < results.length; countActual++, countDisplay++) {

        // Skip if URL is invisible (PAA box, other search features)
        var height = window.getComputedStyle(results[countActual].querySelector('a h3')).height;
        if (height == "auto") { var height = 0; }
        if (height < 20) {
            countDisplay--;
            continue;
        }

        if (currentPage == 1) {
            var count = countDisplay;
        } else if (parseInt(localStorage.getItem('countLastPage')) !== null && (parseInt(localStorage.getItem('lastPage')) + 1) == currentPage) {
            var count = (countDisplay + parseInt(localStorage.getItem('countLastPage')));
        } else {
            // Fallback for when localStorage is not set (no chronological navigation)
            var count = (countActual + (currentPage * resultsPerPage)) - (resultsPerPage-1);
        }

        let counter = document.createElement('div');
        counter.className = 'mst-serp-counter';
        counter.innerHTML = count+'<span>►</span>' ;

        results[countActual].parentNode.parentNode.parentNode.append(counter);
    }

    localStorage.setItem('lastPage', currentPage);
    localStorage.setItem('countLastPage', count);
    return true;
}

// Listen to hash change
// window.onhashchange and similar functions don't work with instant search
oldHash = location.hash;
var hashchange = setInterval(function() {
    currentHash = location.hash;
    if (currentHash != oldHash) {
        // hash has changed, run function again
        setTimeout(main, 750);
        // set curenthash again
        // and keep listening
        oldHash = currentHash;
    }

}, 750);


function ok() {
    const tags = [...document.querySelectorAll('.ok .serpcounter')];
    const texts = new Set(tags.map(x => x.innerHTML));
    tags.forEach(tag => {
        if (texts.has(tag.innerHTML)) {
            texts.delete(tag.innerHTML);
        } else {
            tag.remove()
        }
    });
}
 


// run on page load 
chrome.storage.local.get(['key'], function(result) {
  // console.log('Value currently is ' + result.key);
   if(result.key != 'mst-toggle-off'){
      setTimeout(mstSerpMain, 750);

       
document.addEventListener("scrollend", (event) => {
    mstSerpMain();
});
   }

});




