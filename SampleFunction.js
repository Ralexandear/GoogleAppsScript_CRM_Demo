function sendShuttleList(){

    let shuttleList = getShuttleList();
    let listOfCouriers = getListOfCouriers();
    let userList = Users.getRange('b:f').getValues();
    
    if (! ( shuttleList && listOfCouriers ) ) return;
  
    let message = '';
    let messageEntities = [];
  
    createMessage();
    return sendMessage();
  
  
  
    function createMessage(){
      let shuttleListOfCouriers;
      let userData;
    
      for (let i = 1; i < shuttleList.length; i++ ){
  
        shuttleListOfCouriers = getShuttleListOfCouriers( shuttleList[i] );
        userData = getCourierData( shuttleListOfCouriers );
  
        if (! userData ) continue;
  
        addMessage( userData );
      }
    }
  
  
    function addMessage(userData){
      let text;
      let point;
      let name;
      let telegramUrl;
      let phone_number;
      let offset;
      let nameLength;
      let messageLength;
  
      for ( let i = 0 ; i < userData.length ; i++ ){
  
        point = userData[i][0];
        name = userData[i][1];
        telegramUrl = userData[i][2];
        phone_number = userData[i][3];
        messageLength = message.length;
  
        if ( i == 0 ){
          let adress = shuttleList[i][1];
  
          text = `\n\n🚚🚛 Челнок: ${point+' '+name}\n📞${phone_number}\n🚗 Место встречи: ${adress}\nКурьеры:`;
  
          offset = text.indexOf('Место встречи: ')+15+messageLength;
          
          addGeoLink( offset, adress );
          
          offset = messageLength + point.length + 16;
          if ( messageLength ) offset+2;
  
          nameLength = name.length;
  
        } else {
          text = '\n' + point + ' '+ name + '\n📞' + phone_number;
          offset = messageLength + point.length + 2;
          nameLength = name.length;
        }
  
        addTextLink( offset, nameLength, telegramUrl);
        message = message + text;
      }
  
  
      function addTextLink( offset, length, url){
        messageEntities
          .push( {
            "offset": offset,
            "length": length,
            "type": "text_link",
            "url": url
          } ) ;
      }
  
      function addGeoLink( offset, adress ){
        try{
          let navi = Maps.newGeocoder().geocode( `Санкт-Петербург, ${adress}` );
          let latitude = navi.results[0]['geometry'].location['lat'].toFixed(6);
          let longtitude = navi.results[0]['geometry'].location['lng'].toFixed(6);
  
          let navi_url = `https://yandex.ru/navi/?whatshere%5Bzoom%5D=18&whatshere%5Bpoint%5D=${longtitude}%2C${latitude}&lang=ru&from=navi`;
  
  
          messageEntities
            .push( {
              "offset": offset,
              "length": adress.length,
              "type": "text_link",
              "url": navi_url
            } );
  
        } catch(err) { Logger.log(err) }
      }
    }
  
  
  
    function getShuttleListOfCouriers(shuttleList){
      let courierRow = shuttleList[2].toString().replace(/[,]/g, '.');
      let shuttleListOfCouriers;
  
      if (courierRow.includes(';')) shuttleListOfCouriers = courierRow.split(';'); //Если у челнока несколько курьеров
      else shuttleListOfCouriers = [courierRow.toString()]; //Если только 1
  
      shuttleListOfCouriers.unshift(shuttleList[0]);
      return shuttleListOfCouriers; // Ставим челнока нулевым элементом массива
    }
  
  
    function getCourierData( shuttleListOfCouriers ){
  
      let array = [];
      let point;
      let name;
      let telegramUrl;
      let phone_number;
  
      for ( let u = 0 ; u < shuttleListOfCouriers.length; u++ ){
        if (! shuttleListOfCouriers[u] ) continue;
      
        point = 'К'+shuttleListOfCouriers[u];
        let nameAndPhone = getNameAndPhone( point );
        if (! nameAndPhone ) continue;
       
        name = nameAndPhone[0];
        phone_number = nameAndPhone[1];
        
        if (! name ) continue; //Можно вписать чего делать если точка не найдена
        
        telegramUrl = getTelegramLink( name );
  
        if (! telegramUrl) telegramUrl = null;
        
        array.push([point, name, telegramUrl, phone_number]);
      }
      return array;
  
      
      function getNameAndPhone(point){
          let pointFromList;
  
          for ( let z = 0 ; z < listOfCouriers.length ; z++ ){
            pointFromList = listOfCouriers[z][0];
  
            if ( point !== pointFromList ) continue;
            
            let name = listOfCouriers[z][2];
            let phone_number = listOfCouriers[z][3];
            return [name, phone_number];
          }
      }
  
  
      function getTelegramLink( name ){
        let nameFromList;
        for ( let p = 0 ; p < userList.length ; p++ ){
          nameFromList = userList[p][4];
          
          if ( name === nameFromList ){
            return link = 't.me/'+userList[p][0];
          }
        }
      }
    }
  
    function sendMessage(){
      var data = {
        method: "post",
        payload: {
          method: "sendMessage",
          chat_id: String(newsiD), //iD канала
          text: message,
          entities: JSON.stringify(messageEntities),
          disable_web_page_preview: true
        }
      };
  
      UrlFetchApp.fetch(telegramUrl + '/', data);
    }
    
  
    function getListOfCouriers(){
      try{
        let listOfCouriers = worktable
                                      .getSheetByName( 'График на завтра' )
                                      .getRange( 'A:H' )
                                      .getValues();
                        
        let i = 0;
        while ( listOfCouriers[i][0] ) i++;
  
        listOfCouriers.splice ( i );
        return listOfCouriers;
  
      }
      catch (err) {
        Logger.log( 'getListOfCouriers\n' + err )
      }
    }
    
    function getShuttleList(){
      try{
        let shuttleList = worktable
                                  .getSheetByName( 'Шаблон' )
                                  .getRange( 'I:K' )
                                  .getValues();
  
        let i = 0;                                
        while ( shuttleList[i][0] ) i++;
        
        shuttleList.splice( i );
        return shuttleList;
      }
      catch (err) {
        Logger.log( 'getShuttleList\n' + err )
      }
    }
  }