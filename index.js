const express = require('express');
const app = express();
const bodyParser = require('body-parser');
var reader = require("xlsx");
const nodemailer = require('nodemailer');
app.use(bodyParser.urlencoded({
  extended: true
}));

require('dotenv').config()
app.use(express.static('public'));


//Read excel file
const file = reader.readFile('./Data.xlsx')

let data = []

const sheets = file.SheetNames

for(let i=0;i<sheets.length;i++){
  const temp = reader.utils.sheet_to_json(
  file.Sheets[file.SheetNames[i]])
  temp.forEach((res) => {
    data.push(res.Email)
  })
}

console.log(data)



app.get("/", function(req, res) {
  res.sendFile(__dirname + "./public/index.html")
})

app.post("/",function(req,res){
  const  nme = req.body.email2;
  const mobile = req.body.Mobile;
  const sub=req.body.subject;
  const txt=req.body.message;
  console.log(nme+"\n"+mobile+"\n"+sub+"\n"+txt);
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {  
      user: "bulkemail25@gmail.com",
      pass: "jnyassasoizwljer"
    }
  })
  var myEmails = data
  const options = {
      from:"bulkemail25@gmail.com",
      to:myEmails,
      subject:sub,
      text:txt
  }
  transporter.sendMail(options,function(err,info){
      if(err){
          console.log(err);
          return;
      }
      console.log("Sent "+ info.response);
  })

  const accountSid = process.env.TWILIO_ACCOUNT_SID;
  const authToken = process.env.TWILIO_AUTH_TOKEN;
  const client = require('twilio')(accountSid, authToken);
  
  client.messages 
  .create({body: sub, from: '+14068046817', to: mobile})
  .then(message => console.log(message.sid));
  
  res.sendFile(__dirname + '/success.html')

})

app.listen(process.env.PORT || 5000, function() {
  console.log("Server up and running at port 5000.");
})



