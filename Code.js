MAIL_LABEL_NAME = "sbbol-processed";
ATTACHMENT_SEARCH = 'Выписка за ';
MAIL_SEARCH = '"' + ATTACHMENT_SEARCH + '" has:attachment "zip" -label:' + MAIL_LABEL_NAME;
ENCODING_1C = "Windows-1251";
OUR_SBER_ACC = '40703810856000002997';

const getData = (arr) => {
  const data = arr.map((item) => {
    let o = {}
    let i  = item.split("\r\n")
  
    i.map((item) => {
      if( item.indexOf("=") === -1 ) return;
      const index = item.indexOf("=") 
  
      o[item.slice(0, index)] = item.slice(index + 1)
    })
    return o
  })

  return data
}

function getOrCreateLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if(!label){
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

function scanMail() {
  var processed_label = getOrCreateLabel(MAIL_LABEL_NAME);
  
  var threads = GmailApp.search(MAIL_SEARCH);
  Logger.log(MAIL_SEARCH);
  var sheet = SpreadsheetApp.getActiveSheet();


  var added_set = [];
  var acc_values = [];
  var msgs = GmailApp.getMessagesForThreads(threads.reverse());
  for (var i = 0 ; i < msgs.length; i++) {
    for (var j = 0; j < msgs[i].length; j++) {
      var attachments = msgs[i][j].getAttachments({includeInlineImages:false});
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var attachment_name = attachment.getName();
        if(attachment_name.indexOf(ATTACHMENT_SEARCH)>=0 && attachment_name.indexOf(".zip")>=0){
          Utilities.unzip(attachment).forEach((file) => {
            var file_name = file.getName();
            if(file_name.indexOf("1c")>=0 && file_name.indexOf(".txt")>=0){
              data_1c = file.getDataAsString(ENCODING_1C);
              
              getData(data_1c.split("СекцияДокумент=Платежное поручение")).forEach((sber_map) => {
                if(sber_map['ПлательщикСчет']){
                  var sum = sber_map['Сумма'];
                  var desk = sber_map['НазначениеПлатежа'];
                  var date = sber_map['Дата'].split('.').reverse().join('-');
                  if(sber_map['ПлательщикСчет']==OUR_SBER_ACC){
                    sum = 0 - sum;
                  }
                  var our_id = `${sber_map['Номер']}|${date}|${sum}`;
                  if(!added_set.includes(our_id)){
                    acc_values.push([desk,sum,date,sber_map['Номер']]);
                    added_set.push(our_id);
                  }
                }
              })
            }
          })
        }
      }
    }
    msgs[i][0].getThread().addLabel(processed_label);
  }
  if(acc_values.length){
    sheet.getRange(sheet.getLastRow()+1, 3, acc_values.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(sheet.getLastRow()+1, 1, acc_values.length, 4).setValues(acc_values);
    sheet.getRange(6, 1, sheet.getLastRow(), 4).removeDuplicates().sort(3);
  }

};
