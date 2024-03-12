(function(){
  'use strict';

  let config;
  let settingsDialog;
  const requestUrl = 'https://moodhood-api.livedigital.space/v1/';


  let auth = 'Bearer ';
  let token = "";

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

      function refresh() {
        console.log('refresh() invoked');
        let _userInfo = localStorage.getItem("userInfo");
        token= localStorage.getItem("token");
        if( token && _userInfo ) {
          console.log('if token - true');
         _userInfo =JSON.parse(localStorage.getItem("userInfo"))
          $('#user-info-mail').text(_userInfo.email)
          setPage('u-loged');
          Office.context.mailbox.item.notificationMessages.removeAsync("ActionPerformanceNotification");
        }else{
          console.log('if token - false');
          setPage('u-login');
          const message = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Нужно автозироваться для использования плагина",
            icon: "Icon.80x80",
            persistent: true,
          };
          Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);
        }
      }
      refresh();

      function setPage(_page){
        console.log('setPage',_page,$(_page));
        $('main.u-set').removeClass('active');
        $('.'+_page).addClass('active');
      }

      $('#save-token-btn').on('click', function(e){
        e.stopPropagation();
        // e.preventDefault();
        console.log('svtkn');
        token= $('#token-input').val();

        try {
          $.ajax({
            url: requestUrl+'users/me',
            dataType: 'json',
            headers: {"Authorization": auth+token}
          }).done(function(response){
            // callback(gists);
            console.log(response);
            localStorage.setItem("token",token);
            localStorage.setItem("userInfo",JSON.stringify(response));
            refresh();

          }).fail(function(error){
            console.log("err");
            // callback(null, error);
          });
        } catch (error) {
            
        }
        
      });
      $('#logout-token-btn').on('click', function(e){
        e.stopPropagation();
        console.log('logout-token-btn clicked');
        localStorage.setItem("token","");
        refresh();
      });


      $('#create-in-new-room-btn').on('click', function(e){
        e.stopPropagation();
        console.log('create-in-new-room-btn clicked');

        Office.context.mailbox.item.subject.setAsync("<a href='https://whoer.net' target='_blank'>Some link</a>",
        { coercionType: "html", },
          function (asyncResult) {
            if (asyncResult.status ==
                Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
          }
        );

        // Office.context.mailbox.item.location.setAsync("https://whoer.net",
        //   { coercionType: "html", },
        //   function (asyncResult) {
        //       if (asyncResult.status ==
        //           Office.AsyncResultStatus.Failed) {
        //           write(asyncResult.error.message);
        //       }
        //   }
        // );
        const locations = [
          {
            id: "Contoso",
            type: Office.MailboxEnums.LocationType.Custom,
            phone: '7988'
          }

        ];
        Office.context.mailbox.item.enhancedLocation.addAsync(locations, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Successfully added locations ${JSON.stringify(locations)}`);
          } else {
            console.error(`Failed to add locations. Error message: ${result.error.message}`);
          }
        });        
        Office.context.mailbox.item.body.setAsync("<a href='https://whoer.net' target='_blank'> Мероприятие</a>",
          { coercionType: "html", },
          function (asyncResult) {
              if (asyncResult.status ==
                  Office.AsyncResultStatus.Failed) {
                  write(asyncResult.error.message);
              }
          }
        );


        refresh();
      });

      $('#create-in-recent-room-btn').on('click', function(e){
       

      });      

    //   function setLabelsForEvent() {
    //     Office.context.mailbox.item.body.setAsync(
    //         "Hello world!",
    //         {
    //             coercionType: "html", // Write text as HTML
    //         },

    //         // Callback method to check that setAsync succeeded
    //         function (asyncResult) {
    //             if (asyncResult.status ==
    //                 Office.AsyncResultStatus.Failed) {
    //                 write(asyncResult.error.message);
    //             }
    //         }
    //     );
    // }

      // When insert button is selected, build the content
      // and insert into the body.
      $('#insert-button').on('click', function(){
        const gistId = $('.ms-ListItem.is-selected').val();
        getGist(gistId, function(gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                  {coercionType: Office.CoercionType.Html}, function(result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError('Could not insert gist: ' + result.error.message);
                    }
                });
              } else {
                showError('Could not create insertable content: ' + error);
              }
            });
          } else {
            showError('Could not retrieve gist: ' + error);
          }
        });
      });

      // When the settings icon is selected, open the settings dialog.
      $('#settings-icon').on('click', function(){
        // Display settings dialog.
        let url = new URI('dialog.html').absoluteTo(window.location).toString();
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
        }

        const dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      })
    });
  };

  function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function(gists, error) {
      if (error) {

      } else {
        $('#gist-list').empty();
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('#insert-button').removeAttr('disabled');
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
  }

  function showError(error) {
    $('#not-configured').hide();
    $('#gist-list-container').hide();
    $('#error-display').text(error);
    $('#error-display').show();
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();