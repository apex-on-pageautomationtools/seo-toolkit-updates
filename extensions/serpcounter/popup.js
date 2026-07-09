document.addEventListener('DOMContentLoaded', function () {
    var checkboxss = document.getElementsByClassName('serp');
    chrome.storage.local.get('enabled', function (result) {
        if (result.enabled != null) {
            checkboxss[0].checked = result.enabled;
        }
    });
    checkboxss[0].addEventListener('change', function () {
        console.log(checkboxss[0].checked);
        chrome.storage.local.set({ 'enabled': checkboxss[0].checked }, function () {
            console.log("confirmed");
        });
    });

 
  // if (checkboxss.checked) {
  //        chrome.storage.local.clear(function() {
  //           var error = chrome.runtime.lastError;
  //           if (error) {
  //               console.error(error);
  //           }
  //           // do something more
  //       });
  //       chrome.storage.sync.clear(); // callback is optional
  //    } else{
         
  //       var value = true;
  //         chrome.storage.local.set({key: 'mst-toggle-off'}, function() {
  //             // console.log('Value is set to ' + value);
  //           }); 
  //    }

});

 document.addEventListener('DOMContentLoaded',function(){
     var checkbox = document.getElementsByClassName('serp');
      for(let i=0; i<checkbox.length; i++){ 
        checkbox[i].addEventListener('change', function () {
            serpcount(checkbox[i]);
         });
       }
   
});
 function serpcount(checkbox){
     if (checkbox.checked) {
                 var value = true;
          chrome.storage.local.set({key: 'mst-toggle-off'}, function() {
              // console.log('Value is set to ' + value);
            }); 

     } else{
         var value =false;
         chrome.storage.local.clear(function() {
            var error = chrome.runtime.lastError;
            if (error) {
                console.error(error);
            }
            // do something more
        });
        chrome.storage.sync.clear(); // callback is optional
       
     }
 }
