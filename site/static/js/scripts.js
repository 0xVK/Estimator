function send_form() {

  var msg   = $('#main_form').serialize();
  var content_block = $('#content-block');
  var loader_block = $('#loader-wrapper');

  loader_block.slideDown('slow');
  content_block.slideUp('slow');

  $.ajax({
          type: 'POST',
          url: '/ajax_estimate',
          data: msg,
          success: function(responce) {

            content_block.html(responce);
            loader_block.slideUp('fast');
            content_block.slideDown('slow');

          },
          error:  function(xhr, str){
	             alert('Request error: ' + xhr.responseCode);
          }
        });
    }
