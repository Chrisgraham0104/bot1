var restify = require('restify'); 
var builder = require('botbuilder');  
// Setup Restify Server 
var msgArr= [];
var server = restify.createServer(); 
server.listen(process.env.port || process.env.PORT || 3978, 
function () {    
    console.log('%s listening to %s', server.name, server.url);  
});  
// chat connector for communicating with the Bot Framework Service 
var connector = new builder.ChatConnector({     
    /*appId: process.env.MICROSOFT_APP_ID,     
    appPassword: process.env.MICROSOFT_APP_PASSWORD*/
    appId: process.env.MICROSOFT_APP_ID || 'dd0a1b2d-d32f-46dc-a668-c9e848f84b12',     
    appPassword: process.env.MICROSOFT_APP_PASSWORD || 'hFFQ488#+#okechbJRPY51~' 
});
// Listen for messages from users  
server.post('/api/messages', connector.listen());  

var inMemoryStorage = new builder.MemoryBotStorage();

var connectionJson = [];
var dbJson = [];
var branchMail,branch,query,skey,customerid,country;
var param = [];
// Receive messages from the user and respond by echoing each message back (prefixed with 'You said:') 
var bot = new builder.UniversalBot(connector, [
    function (session){
        session.beginDialog('askCountry');
    },
    function (session,results){
        country = session.message.text.toLowerCase();
        //session.beginDialog('askBranch');
        session.beginDialog('askKey');
    },
    function (session, results) {     
       
        //session.send('New Session');
        console.log('results ---> ' + JSON.stringify(results));
        //session.send('Your Key value is : ' +  results.response);
        var responseText = session.message.text.toLowerCase();
        skey = responseText;
        console.log(responseText);
        var flipflop = 0;
        for (var i=0; i<dbJson.length;i++){
            var regex = dbJson[i].Key.toLowerCase();
            console.log('regex -->' + regex + ' Text -->' + responseText);
            var strCount = responseText.search(regex);
            //session.send(strCount);
            console.log('count --->' + strCount)
            if(strCount >= 0){
                query = dbJson[i].Query;
                param = dbJson[i].Prameters.split(",");
                console.log('Query ---> '+ query);
                console.log('Param --> ' +  JSON.stringify(param));
                flipflop=1;
                break;
            }

        }

        if(flipflop == 0){
            session.send('No Keyword Found');
        }else {
      
             //session.beginDialog('askQuery');        
             session.beginDialog('askBranch');
        }   

    },
    function (session, results){
         //session.send("Good morning.");
        //session.send(session.message);
        console.log('results ---> ' + JSON.stringify(results));
        var responseText = session.message.text.toLowerCase();
        customerid = responseText;
        //session.send(responseText.toLowerCase());
        var flipflop = 1; //if db purpose set 0
        /*for (var i=0; i<connectionJson.length;i++){
            regex = connectionJson[i].BranchID.toLowerCase();
            console.log('regex -->' + regex + ' Text -->' + responseText);
            var strCount = responseText.search(regex);
            //session.send(strCount);
            console.log('count --->' + strCount)
            if(strCount >= 0){
                branchMail = connectionJson[i].BranchMail;
                branch = connectionJson[i].BranchID;
                session.send('You selected Branch: ' + branch + ' & Branch Mail: ' + branchMail);
                flipflop=1;
                break;
            }

        }*/

        var result = encodeURIComponent(query);
        var rtStr = getFunction(result);
        //var geturl = "http://localhost:8083/query/" + country + "-" + results.response;
        var geturl = "https://ipfapi.azurewebsites.net/query/" + country + "-" + results.response;        
        console.log('CNFS query url --> ' + geturl);
        var request = require('request');

        request(geturl, function (error, response, body) {
            console.log('----------------------API CALL-------------------------')
            //session.send('into api call')    ;
            console.log('error:', error); // Print the error if one occurred
            console.log('statusCode:', response && response.statusCode); // Print the response status code if a response was received
            console.log('body:', body); // Print the HTML for the Google homepage.
            console.log(JSON.parse(body).recordset);
            body=JSON.parse(body);
            if(error){
                sessions.send('Error Occurs Err: ' + error);
            }else{
                if(response.statusCode == 200){
                    if((body.recordset[0] !== undefined)){
                        /*keys = Object.keys(body.recordset[0]);
                        template = 'Result of Search : <br/> ';
                        _ref = body.recordset[0];
                        for (k in _ref) {
							v = _ref[k];
							template = template + k + ' : ' + v + '<br/>';
                        } */  
                        flipflop=1;
                        branch = 'LN'+body.recordset[0].Agency_Segment_Db_Id;
                        //branchMail = connectionJson[i].BranchMail;
                        console.log('New Branch :' + branch);
                        session.beginDialog('askQuery');
                    }else{
                        flipflop=0;
                        session.send('No Database Found');
                    }
                }else{
                    sessions.send('Response Code : ' + response.statusCode + "<br/>"+ 'Response : ' + response);
                }
            }
        });        
        console.log('----------------------After API CALL-------------------------')
       /* if(flipflop == 0){
            session.send('No Database Found');
        }else {           


           // session.beginDialog('askKey');        
           session.beginDialog('askQuery');
        }*/   
   
       
    },
    function (session,results){
        console.log('results ---> ' + JSON.stringify(results));
       // session.send('Your Key value is : ' +  results.response);
        var val = results.response.split(' ');
        var exeQuery = makeQuery(val,query);

       // session.send('Your Execution query is : ' + exeQuery);
        query = exeQuery;
        session.send('You are Searching for "' + capitalize(country) + " " + capitalize(customerid) + " " + capitalize(skey) + " " + results.response + '"');

        session.beginDialog('askExecution');

    },
    function (session,results){
        //session.send(`Reservation confirmed. Reservation details: <br/>Date/Time: ${session.dialogData.reservationDate} <br/>Party size: ${session.dialogData.partySize} <br/>Reservation name: ${session.dialogData.reservationName}`);
        console.log('Bye hit');
        session.endDialog();
        session.reset();        
    }
]).set('storage', inMemoryStorage); // Register in-memory storage 

// Dialog to ask for a Key query
bot.dialog('askKey', [
    function (session) {
        
        var fpath = '.\\xls' + '\\configData.xlsx';
        configDataArr = configFile(fpath, 'Type', 'query');
        var dbfilepath = '.\\' + configDataArr.Folder + '\\' + configDataArr.Filename;
        var dbfile = require('xlsx').readFile(dbfilepath);
    
        dbJson = xlstojson(dbfilepath);
        console.log ('db count data ' + dbJson.length);

        var queryExe = '';
        var queryStr = '';
        for(var i = 0, j = 1; i < dbJson.length; i++, j++){
            queryExe += j + "] " + dbJson[i].Key + "<br/>";
            queryStr += dbJson[i].Key;
            if(j<dbJson.length){
                queryStr += '|';
            }            
        }    

                  
            var msg = "Select value from below list for further execution <br/>" + queryExe;
            /*  var reply = new builder.Message()    
                    .address(message.address)    
                    .text(msg);   */
           // session.send(msg);    
       // builder.Prompts.text(session, msg);
        builder.Prompts.choice(session, "Select value from below list for further execution <br/>", queryStr, { listStyle: builder.ListStyle.button });
    },
    function (session, results) {
        session.endDialogWithResult(results);
    }
]);

bot.dialog('askQuery', [
    function (session) {
        var paramstr='';
        for(var i=0, j=1; i<param.length; i++,j++){
            paramstr += '[ ' + param[i] + ' ]';
            if(j<param.length){
                paramstr += ' ';
            }
        }
        var msg = 'Request you to provide Below Parameter for the Search <br/>' + paramstr;       
        builder.Prompts.text(session,msg);
        
    },
    function (session, results) {
        session.endDialogWithResult(results);
    }
]);

bot.dialog('askExecution', [
    function (session) {
        
        var result = encodeURIComponent(query);
        var rtStr = getFunction(result);
        //var geturl = "http://localhost:8083/query/" + country + "-" + branch + "/" + rtStr;
        var geturl = "https://ipfapi.azurewebsites.net/query/" + country + "-" + branch + "/" + rtStr;
        console.log('query url --> ' + geturl);
        var request = require('request');

        request(geturl, function (error, response, body) {
            //session.send('into api call')    ;
            console.log('error:', error); // Print the error if one occurred
            console.log('statusCode:', response && response.statusCode); // Print the response status code if a response was received
            console.log('body:', body); // Print the HTML for the Google homepage.
            console.log(JSON.parse(body).recordset);
            body=JSON.parse(body);
            if(error){
                //sessions.send('Error Occurs Err: ' + error);
                msg = 'Error Occurs Err: ' + error + "<br/>" + 'Type "Bye" or "1" To Exit from current session';
                builder.Prompts.text(session,msg);
            }else{
                if(response.statusCode == 200){
                    if((body.recordset[0] !== undefined)){
                        keys = Object.keys(body.recordset[0]);
                        template = 'Result of Search : <br/> ';
                        _ref = body.recordset[0];
                        for (k in _ref) {
                        v = _ref[k];
                        template = template + k + ' : ' + v + '<br/>';
                        }   

                        /* generating excel file*/
                        fs = require('fs');
                        json2xls = require("json2xls");
                        json = body.recordset[0];
                        xls = json2xls(json);
                        //configDataArr = configFile(fpath, 'Type', 'outputxls');
                        /* set date format*/

                        today = new Date;
                        hour = today.getHours();
                        minute = today.getMinutes();
                        second = today.getSeconds();
                        dd = today.getDate();
                        mm = today.getMonth() + 1;
                        yyyy = today.getFullYear();
                        fdate = '_' + dd + mm + yyyy + '-' + hour + minute + second;
                        /* end Date format*/
                        //shortCountry =  body.recordset[0].countrypseudo;

                        fname = branch + '_' + customerid + fdate + '.xlsx';
                        ///xfilepath = '.\\' + configDataArr.Folder + '\\' + fname;
                        //xfilepath = '.\\' + 'xls_data' + '\\' + fname;
			    xfilepath = './' + 'xls_data' + '/' + fname;
                        console.log('xls file path --> ' + xfilepath);
                        //fs.writeFileSync(xfilepath, xls, 'binary');
                        /* azure storage call */
                        var azure = require('azure-storage');
                        var accountName = "demofinbe0c";
                        var accessKey = "ws4xfWVmdNG574lq6CxZJN8/DSfkD9d5zd4CK/dM9YKOUC+C570eLpZJxYFR4ehGooOp1KEhhfwWRM9p+rlnqQ==";
                        var host = "https://demofinbe0c.blob.core.windows.net/";
                        var blobService = azure.createBlobService(accountName, accessKey, host);

                        blobService.createBlockBlobFromText('xls-data', fname, xls,  function(error, result, response){
                            if (error) {
                                console.log('Upload filed, open browser console for more detailed info.');
                                console.log(error);
                            } else {
                                   console.log('Upload successfully!');

                                 /* mail Function*/
                                    if((body.recordset[1] !== undefined)){
                                        branchMail = body.recordset[1][0].email;
                                    }   
                                    console.log('Sender Mail ---> ' + branchMail);
                                    attach = [];
				    //xfilepath = 'https://demofinbe0c.blob.core.windows.net/xls-data/' + fname
				    xfilepath = 'https://demofinbe0c.blob.core.windows.net/xls-data/';
                                    tmpAtt = {
                                    	filename: fname,
					path: xfilepath                                    	
                                    };
                                    attach.push(tmpAtt);
                                    console.log('Attache --> ' + JSON.stringify(attach));
                                    nodemailer = require("nodemailer");
                                    smtpTransport = nodemailer.createTransport({
                                    /*service: "Gmail",
				    host: "smtp.gmail.com",*/
				    host: "smtp-mail.outlook.com", // hostname
				    secureConnection: false, // TLS requires secureConnection to be false
				    port: 587, // port for secure SMTP
				    tls: {
				       ciphers:'SSLv3'
				    },
                                    auth: {
                                        user: 'hubotest23@outlook.com',
                                        pass: 'S$35@v$#@_'
                                    }
					  
                                    });
                                    mail = {
                                        from: "hubotest23@outlook.com",
                                        to: branchMail,
                                        subject: branch + ' ' + customerid + ' Data File',
                                        text: "Auto Generated Mail Contain Response Data",
                                        attachments: attach                                        
                                    };
                                    smtpTransport.sendMail(mail, function(error, response) {
                                    if (error) {
                                        console.log(error);
                                    } else {
                                        console.log("Message sent: " + response.response);
                                    }
                                    return smtpTransport.close();
                                    });
                                    /* end Mail function*/
                                        
                                    //session.send(template);
                                    //session.send('Type "Bye" To Exit from current session');
                                    msg = template + "<br/>" + 'Type "Bye" or "1" To Exit from current session';
                                //builder.Prompts.text(session,msg);
                                builder.Prompts.choice(session, msg, ["Bye"]);
                                    //session.endDialog(); 

                            }
                        });


                        /* end storage call */


                        /* end generating excel file*/
                        
                    }else{
                        var noStr = 'Kindly update Search Parameters';
                        msg = noStr + "<br/>" + 'Type "Bye" or "1" To Exit from current session';
                        builder.Prompts.choice(session, msg, ["Bye"]);
                        //builder.Prompts.text(session,msg);
                    }
                }else{
                    //sessions.send('Response Code : ' + response.statusCode + "<br/>"+ 'Response : ' + response);
                    msg = 'Response Code : ' + response.statusCode + "<br/>"+ 'Response : ' + capitalize(body.error) + "<br/>" + 'Type "Bye" or "1" To Exit from current session';
                    builder.Prompts.text(session,msg);
                }
            }
        });
        
    },
    function (session, results) {
        session.endDialogWithResult(results);
    }
]);

bot.dialog('askBranch',[
    function (session){
        //msgArr = message;
        var fpath = '.\\xls' + '\\configData.xlsx';
        configDataArr = configFile(fpath, 'Type', 'connection');
        console.log('Connection config --> ' + configDataArr.Folder);
        var connfilepath = '.\\' + configDataArr.Folder + '\\' + configDataArr.Filename;
        var connfile = require('xlsx').readFile(connfilepath);
    
        connectionJson = xlstojson(connfilepath);
        console.log ('count data ' + connectionJson.length);
    
        var branchData = '';
        var branchStr = '';
        for(var i = 0, j = 1; i < connectionJson.length; i++, j++){
                branchData += j + "] " + connectionJson[i].BranchID + "<br/>";
                branchStr += connectionJson[i].BranchID;
                if(j<connectionJson.length){
                    branchStr += '|';
                }
        }

        //msg ="Kindly Select a Branch from Provided Options <br/>";

       // builder.Prompts.choice(session, msg, branchStr, { listStyle: builder.ListStyle.button });

      msg = 'Kindly Provide Customer Id <br/>';
       builder.Prompts.text(session,msg);
    }
]);

bot.dialog('askCountry',[
    function (session){
        

      msg = 'Kindly Provide Country Name <br/>';
      builder.Prompts.choice(session, msg, ["Poland","Romania","Czech","Mexico","Hungary","Lithuania"], { listStyle: builder.ListStyle.button });
    }
]);
bot.on('conversationUpdate', function (message) {
    /*msgArr = message;
    var fpath = '.\\xls' + '\\configData.xlsx';
    configDataArr = configFile(fpath, 'Type', 'connection');
    console.log('Connection config --> ' + configDataArr.Folder);
    var connfilepath = '.\\' + configDataArr.Folder + '\\' + configDataArr.Filename;
    var connfile = require('xlsx').readFile(connfilepath);

    connectionJson = xlstojson(connfilepath);
    console.log ('count data ' + connectionJson.length);

    var branchData = '';
    for(var i = 0, j = 1; i < connectionJson.length; i++, j++){
            branchData += j + "] " + connectionJson[i].BranchID + "<br/>";
    }
    */
    // console.log (connfile);
    
    if (message.membersAdded[0].id === message.address.bot.id) {
       // var msg = "Hello, Welcome to Chat System of the Demo Application <br/> Kindly Type a Branch from Provided Options <br/>" + branchData;
       var msg ="Hello, Welcome to Chat System  <br/>";
          var reply = new builder.Message()    
                .address(message.address)    
                .text(msg);        
          bot.send(reply);      
          bot.beginDialog(message.address, '/');
    }
   //message.response.send('Hi');
    
   
 }); 

 
function xlstojson(path){
    var XLSX = require('xlsx');
    var workbook = XLSX.readFile(path);
    var sheet_name_list = workbook.SheetNames;
    var data = [];
    sheet_name_list.forEach(function(y) {
        var worksheet = workbook.Sheets[y];
        var headers = {};
        
        for(z in worksheet) {
            if(z[0] === '!') continue;
            //parse out the column, row, and value
            var col = z.substring(0,1);
            var row = parseInt(z.substring(1));
            var value = worksheet[z].v;
    
            //store header names
            if(row == 1) {
                headers[col] = value;
                continue;
            }
    
            if(!data[row]) data[row]={};
            data[row][headers[col]] = value;
        }
        //drop those first two rows which are empty
        data.shift();
        data.shift();
        console.log(data);
        
    });
    
    return data;
}
 function getFunction(Options) {
    var chars,
      _this = this;
    chars = {
      "'": "%27",
      "(": "%28",
      ")": "%29",
      "*": "%2A",
      "!": "%21",
      "~": "%7E"
    };
    return Options.replace(/[~!*()']/g, function(m) {
      return chars[m];
    });
  }

 function makeQuery(data, changeStr){
    var cnt, refstr;
    console.log('old : ' + changeStr);
    cnt = 0;
    while (cnt < data.length) {
      if (cnt > 0) {
        refstr = '{{ref_text' + cnt + '}}';
      } else {
        refstr = '{{ref_text}}';
      }
      changeStr = changeStr.replace(refstr, data[cnt]);
      cnt++;
    }
    console.log('new : ' + changeStr);
    return changeStr;
  }
function configFile(filepath, field, fvalue) {
        var configJson = xlstojson(filepath);

        var xlsData = configJson.find(function(x) {
        if (field === 'Type') {
          return x.Type === fvalue;
        }
        if (field === 'Branch') {
          return x.BranchID === fvalue;
        }
      });
      console.log('xls ---->' + JSON.stringify(xlsData));
      return xlsData;
}

function capitalize(string) {
    return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}
