const fs=require('fs');let captured;
global.Utilities={formatDate:(d)=>d.toDateString()};global.MailApp={sendEmail:o=>captured=o};
eval(fs.readFileSync(require('path').join(__dirname,'..','apps-script','Code.gs'),'utf8'));
const p=JSON.parse(fs.readFileSync(process.argv[2]));
sendReport_(p,'https://docs.google.com/spreadsheets/d/x');
console.log(captured.subject);fs.writeFileSync(process.argv[3],captured.htmlBody);
