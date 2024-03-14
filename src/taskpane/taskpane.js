(function(){
  'use strict';

  let config;
  let settingsDialog;
  const requestUrl = 'https://moodhood-api.livedigital.space/v1/';


  let auth = 'Bearer ';
  let token = "";
  let headers = {
    
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers":"*",
    "Access-Control-Allow-Methods": "*",
    // "Origin": "https://ld-outlook-addin.onrender.com/",
    "Content-Type": "application/json",
    "access-control-allow-credentials" : "true" ,
    "vary": "Origin"
  };
    

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
            cors: true ,
            secure: true,
            headers: {
              ...headers,
              "Authorization": auth + token,
            },
          }).done(function(response){
            // callback(gists);
            console.log(response);
            localStorage.setItem("token",token);
            localStorage.setItem("userInfo",JSON.stringify(response));
            refresh();
            $.ajax({
              url: requestUrl+'users/me/settings/meetings',
              dataType: 'json',
              cors: true ,
              secure: true,              
              headers: {
                ...headers,
                "Authorization": auth + token,
              },
            }).done(function(response){
              // callback(gists);
              console.log(response);
              localStorage.setItem("roomid",response.roomId);
              localStorage.setItem("spaceid",response.spaceId);
            });

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
        let _event_info = {};
        let _roomid = localStorage.getItem("roomid");
        let _spaceid =localStorage.getItem("spaceid")
        token= localStorage.getItem("token");
        console.log('roomid/spaceid = ', _roomid + '/' + _spaceid);
        Office.context.mailbox.item.start.getAsync((result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error(`Action failed with message ${result.error.message}`);
              return;
            }
            console.log(`Appointment starts: ${result.value.getUTCDay()}-${result.value.getUTCMonth()+1}-${result.value.getUTCFullYear()} `);

            _event_info = {
              "name": `Встреча Outlook ${result.value.getUTCDate()}-${result.value.getUTCMonth()+1}-${result.value.getUTCFullYear()}`,
              "isPublic": true,
              "isScreensharingAllowed": true,
              "isChatAllowed": true,
              "type": "lesson"
            };


            try {
              $.ajax({
                url: requestUrl+'spaces/'+_spaceid+'/rooms',
                method: "POST",
                cors: true ,
                secure: true,                
                headers: {
                  ...headers,
                  "Authorization": auth + token,
                },
                data:JSON.stringify(_event_info)
              }).done(function(response){
                // callback(gists);
                console.log(response);
                setLabelsForEvent(response.name,response.alias);

                Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", 
                {
                  type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                  message: "Мероприятие добавлено",
                  icon: "Icon.80x80",
                  persistent: true,
                });
                
                // refresh();
              }).fail(function(error){
                console.log("err",error);
                // callback(null, error);
              });
            } catch (error) {
                
            }
        });  
        // refresh();
      });

      $('#create-in-recent-room-btn').on('click',  function(e){
        let val = '';
        console.log('val=', val);
        if (!$('#custom-room-block').hasClass('closed')) {
          $('#custom-room-block').addClass('closed')
        } else {
          let _spaceid = localStorage.getItem("spaceid")
          $('#spaces-select')[0].innerHTML = "";
          // fabricRefresh();
          $.ajax({
            url: requestUrl + 'spaces/',
            method: "GET",
            cors: true ,
            secure: true,            
            headers: {
              ...headers,
              "Authorization": auth + token,
            },
          }).done(function (response) {
            console.log(response);
            let _optionsList = ''
            let _spacesArray = response.items;
            for (let i = 0; i < _spacesArray.length; i++) {
              _optionsList += `<option value="${_spacesArray[i].id}">${_spacesArray[i].name}</option>`;
            }
            console.log("_optionsList", _optionsList, $('#spaces-select')[0]);
            $('#spaces-select')[0].innerHTML = _optionsList;

            $('#custom-room-block').removeClass('closed');
          }).fail(function (error) {
            console.log("err", error);
          });

          $.ajax({
            url: requestUrl + 'spaces/' + _spaceid + '/rooms',
            method: "GET",
            cors: true ,
            secure: true,            
            headers: {
              ...headers,
              "Authorization": auth + token,
            },
          }).done(function (response) {
            console.log(response);
            let _optionsList = ''
            let _roomsArray = response.items;
            for (let i = 0; i < _roomsArray.length; i++) {
              _optionsList += `<option value="${_roomsArray[i].id}" room-alias="${_roomsArray[i].alias}">
              ${_roomsArray[i].name}</option>`;
            }
            console.log("_optionsList", _optionsList, $('#rooms-select')[0]);
            $('#rooms-select')[0].innerHTML = _optionsList;
          }).fail(function (error) {
            console.log("err", error);
          });
        }
      });
      

      $('#choose-in-recent-room-btn').on('click', function (e) {
        let _subject = $("#rooms-select :selected").text().trim();
        let _alias = $("#rooms-select :selected").attr('room-alias');
        setLabelsForEvent(_subject, _alias);
        let _space_settings = {}
        _space_settings['spaceId'] = $('#spaces-select').val();
        _space_settings['roomId'] = $('#rooms-select').val();
        $.ajax({
          url: requestUrl+'users/me/settings/meetings',
          method: "PUT",
          cors: true ,
          secure: true,          
          headers: {
            ...headers,
            "Authorization": auth + token,
          },
          data:JSON.stringify(_space_settings)
        }).done(function(response){
          console.log(response);
          // refresh();
        }).fail(function(error){
          console.log("err",error);
          // callback(null, error);
        });         

      });
      $('#spaces-select').on('change', function (e) {
        let val = '';
        console.log('val=', val);
        let _spaceid = $('#spaces-select').val();
        $.ajax({
          url: requestUrl + 'spaces/' + _spaceid + '/rooms',
          method: "GET",
          cors: true ,
          secure: true,          
          headers: {
            ...headers,
            "Authorization": auth + token,
          },
        }).done(function (response) {
          console.log(response);
          let _optionsList = ''
          let _roomsArray = response.items;
          for (let i = 0; i < _roomsArray.length; i++) {
            _optionsList += `<option value="${_roomsArray[i].id}" room-alias="${_roomsArray[i].alias}">
            ${_roomsArray[i].name}</option>`;
          }
          console.log("_optionsList", _optionsList, $('#rooms-select')[0]);
          $('#rooms-select')[0].innerHTML = _optionsList;
        }).fail(function (error) {
          console.log("err", error);
        });

       
        
      });

      function setLabelsForEvent(_subject,_alias) {
        let body = `<a href="https://edu.livedigital.space/room/${_alias}" target="_blank"> Ссылка на мероприяте ${_subject}</a>`
        Office.context.mailbox.item.subject.setAsync(_subject,
        { coercionType: "html", },
          function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
          }
        );

 
        Office.context.mailbox.item.location.setAsync('https://edu.livedigital.space/room/'+_alias, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Successfully added _alias ${JSON.stringify(_alias)}`);
          } else {
            console.error(`Failed to add locations. Error message: ${result.error.message}`);
          }
        });        
        Office.context.mailbox.item.body.setAsync(body,
          { coercionType: "html", },
          function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                  write(asyncResult.error.message);
              }
          }
        );
      }



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
    }); //jQuery(document).ready
  }; //Office.initialize

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