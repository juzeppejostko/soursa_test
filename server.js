const e = require('express');
const express = require('express');
const { link, writeFile} = require('fs');
const app = express();
const port = 5500;
const path = require('path');

const graphConfig = {
    graphEndpoint: 'https://graph.microsoft.com/v1.0',
    scopes: ['user.read', 'mail.read'] // Разрешения, необходимые для чтения пользовательских данных и электронной почты
};

function formatDate(dateString) {
    const date = new Date(dateString);
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear().toString();
  
    return `${day}-${month}-${year}`;
  }

async function getAllInfo(accessToken, startDate, endDate)
{   

    const firstLinks = {"sentMails": `${graphConfig.graphEndpoint}/me/mailFolders/SentItems/messages?$top=100&filter=receivedDateTime ge ${startDate} and receivedDateTime le ${endDate}`,
                        "receivedMails": `${graphConfig.graphEndpoint}/me/mailFolders/inbox/messages?$top=100&filter=receivedDateTime ge ${startDate} and receivedDateTime le ${endDate}`
                       }
    
    let link
    let MailsInfo = {}
    for (let key in firstLinks) 
    {   
        link = firstLinks[key]
        MailsInfo[key] = []
        while (link !== '') {
            try {
              const response = await fetch(link, {
                headers: {
                  Authorization: `Bearer ${accessToken}`
                }
              });
              
              const data = await response.json();
              if ('@odata.nextLink' in data) {
                link = data['@odata.nextLink'];
              } else {
                link = '';
              }
              for (let i=0; i<data.value.length; i++)
              {
                MailsInfo[key].push(data.value[i])
              }
            } catch (error) {
              console.error('Ошибка при вызове MS Graph API:', error);
              link = ''
              
            }
        }
    }
    
    
    return MailsInfo

}

function getSentMails(MailsInfo)
{
    var sentMailsInfo = {}
    sentMailsInfo.dates = []
    sentMailsInfo.messages = []
    MailsInfo.forEach(message => {
        var date = formatDate(message['sentDateTime'])
        if (sentMailsInfo.dates.includes(date))
        {
            sentMailsInfo.messages[sentMailsInfo.dates.indexOf(date)] += 1
        }
        else
        {
            sentMailsInfo.dates.push(date)
            sentMailsInfo.messages.push(1)
        }
      });
    return sentMailsInfo
}
function getReceivedMails(MailsInfo)
{
    var receivedMailsInfo = {}
    receivedMailsInfo.dates = []
    receivedMailsInfo.messages = []
    MailsInfo.forEach(message => {
        var date = formatDate(message['receivedDateTime'])
        if (receivedMailsInfo.dates.includes(date))
        {
            receivedMailsInfo.messages[receivedMailsInfo.dates.indexOf(date)] += 1
        }
        else
        {
            receivedMailsInfo.dates.push(date)
            receivedMailsInfo.messages.push(1)
        }
      });
    return receivedMailsInfo
}

function convertSeconds(seconds) {
    const days = Math.floor(seconds / (3600 * 24));
    const hours = Math.floor((seconds % (3600 * 24)) / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const remainingSeconds = seconds % 60;

    return {
      days: days,
      hours: hours,
      minutes: minutes,
      seconds: remainingSeconds
    };
}

function getReplyToSent(sentMails, receivedMails, type1, type2)
{
    var deltaArray = []
    for (let i = 0; i<sentMails.length; i++)
    {
        var conversationId = sentMails[i]['conversationId']
        var sentTime = sentMails[i][type1]
        var unixSentTime = Math.floor(Date.parse(sentTime) / 1000);
        var minTime = 0
        var receivedTime = 0
        var unixReceivedTime = 0
        for (let j = 0; j<receivedMails.length; j++)
        {
            if (conversationId == receivedMails[j]['conversationId'])
            {
                receivedTime = receivedMails[j][type2]
                unixReceivedTime = Math.floor(Date.parse(receivedTime) / 1000)
                if (unixReceivedTime > unixSentTime)
                {
                    if (minTime == 0)
                    {
                        minTime = unixReceivedTime
                    }
                    else
                    {
                        if (unixReceivedTime < minTime)
                        {
                            minTime = unixReceivedTime
                        }
                    }
                }
            }
        }
        if (minTime != 0){
            deltaArray.push(minTime - unixSentTime)
        }
    }
    var sum = 0
    for (let i = 0; i < deltaArray.length; i++)
    {
        sum += deltaArray[i]
    }
    average = sum / deltaArray.length
    average = Math.trunc(average)
    return convertSeconds(average)
}

function getDayOfWeek(unixTime) {
    const daysOfWeek = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
    const date = new Date(unixTime * 1000);
    const dayOfWeek = date.getDay();
  
    return daysOfWeek[dayOfWeek];
}
function getHourFromUnixTime(unixTime) {
    // Создание объекта Date с использованием Unix времени (в миллисекундах)
    const date = new Date(unixTime * 1000);
  
    // Получение часа из объекта Date
    const hour = date.getHours();
  
    return hour;
  }

function getDayMailAnalytics(MailsInfo, type)
{
    days = ['Вс', 'Пн', 'Вт', 'Ср', "Чт", "Пт", "Сб"]
    messageData = []
    for (let i=0; i<7; i++)
    {   
        messageData.push({dayOfWeek: days[i], hour: 0, count: 0})
        for (let j=1; j<24; j++)
        {
            messageData.push({dayOfWeek: days[i], hour: j, count: 0})
        }
    }
    
    for (let i=0; i<MailsInfo.length; i++)
    {
        sentTime = MailsInfo[i][type]
        unixSentTime = Math.floor(Date.parse(sentTime) / 1000);
        day = getDayOfWeek(unixSentTime)
        hour = getHourFromUnixTime(unixSentTime)
        for (let j=0; j<messageData.length; j++)
        {
            if (messageData[j].dayOfWeek == day && messageData[j].hour == hour)
            {
                messageData[j].count++
            }
        }
    }
    return messageData
}

function getTopSenders(MailsInfo, top)
{
    senders = {}
    for (let i = 0; i<MailsInfo.length; i++)
    {
        if (MailsInfo[i]['sender']['emailAddress']['address'] in senders)
        {
            senders[MailsInfo[i]['sender']['emailAddress']['address']]['received'] += 1
        }
        else
        {
            senders[MailsInfo[i]['sender']['emailAddress']['address']] = {'received': 1, 'sent': 0}
        }
    }
    const entries = Object.entries(senders);

    entries.sort((a, b) => b[1].received - a[1].received);

    topSenders = Object.fromEntries(entries);

    topSenders = getTopTen(topSenders)

    for (let key in topSenders)
    {
        if (key in top)
        {
            topSenders[key]['sent'] = top[key]['sent']
        }
    }
    return topSenders
}

function getTopRecipients(MailsInfoSent, MailsInfoReceived)
{
    recipients = {}
    for (let i = 0; i<MailsInfoSent.length; i++)
    {
       for (let j = 0; j<MailsInfoSent[i]['toRecipients'].length; j++)
       {
            if (MailsInfoSent[i]['toRecipients'][j]['emailAddress']['address'] in recipients)
            {
                recipients[MailsInfoSent[i]['toRecipients'][j]['emailAddress']['address']]['sent'] += 1
            }
            else
            {
                recipients[MailsInfoSent[i]['toRecipients'][j]['emailAddress']['address']] = {'sent': 1, 'received': 0}
            }
       } 
    }
    const entries = Object.entries(recipients);

    entries.sort((a, b) => b[1].sent - a[1].sent);

    topRecipients = Object.fromEntries(entries);
    
    

    for (let i = 0; i<MailsInfoReceived.length; i++)
    {
        if (MailsInfoReceived[i]['sender']['emailAddress']['address'] in topRecipients)
        {
            recipients[MailsInfoReceived[i]['sender']['emailAddress']['address']]['received'] += 1
        }
    }
    return topRecipients
}

function getReplyToReceiveByTop(sentMails, receivedMails, type1, type2, top)
{      
    for (let key in top)
    {
        var deltaArray = []
        for (let i = 0; i<sentMails.length; i++)
        {
            var conversationId = sentMails[i]['conversationId']
            var sentTime = sentMails[i][type1]
            var unixSentTime = Math.floor(Date.parse(sentTime) / 1000);
            var minTime = 0
            var receivedTime = 0
            var unixReceivedTime = 0
            
            if (sentMails[i]['sender']['emailAddress']['address'] !== key)
            {
                continue
            }
           
            for (let j = 0; j<receivedMails.length; j++)
            {
                if (conversationId == receivedMails[j]['conversationId'])
                {
                    receivedTime = receivedMails[j][type2]
                    unixReceivedTime = Math.floor(Date.parse(receivedTime) / 1000)
                    if (unixReceivedTime > unixSentTime)
                    {
                        if (minTime == 0)
                        {
                            minTime = unixReceivedTime
                        }
                        else
                        {
                            if (unixReceivedTime < minTime)
                            {
                                minTime = unixReceivedTime
                            }
                        }
                    }
                }
            }
            if (minTime != 0){
                deltaArray.push(minTime - unixSentTime)
            }
        }
        var sum = 0
        for (let i = 0; i < deltaArray.length; i++)
        {
            sum += deltaArray[i]
        }
        if (deltaArray.length !== 0)
        {
            average = sum / deltaArray.length
            average = Math.trunc(average)
        }
        else
        {
            average = 0
        }
       
        top[key]['average'] = average
    }   
    return top
}

function getReplyToSentByTop(sentMails, receivedMails, type1, type2, top)
{      
    for (let key in top)
    {
        var deltaArray = []
        for (let i = 0; i<sentMails.length; i++)
        {
            var conversationId = sentMails[i]['conversationId']
            var sentTime = sentMails[i][type1]
            var unixSentTime = Math.floor(Date.parse(sentTime) / 1000);
            var minTime = 0
            var receivedTime = 0
            var unixReceivedTime = 0
            
            if (sentMails[i]['toRecipients'][0]['emailAddress']['address'] !== key)
            {
                continue
            }
           
            for (let j = 0; j<receivedMails.length; j++)
            {
                if (receivedMails[j]['sender']['emailAddress']['address'] !== key)
                {
                    continue
                }
                if (conversationId == receivedMails[j]['conversationId'])
                {
                    receivedTime = receivedMails[j][type2]
                    unixReceivedTime = Math.floor(Date.parse(receivedTime) / 1000)
                    if (unixReceivedTime > unixSentTime)
                    {
                        if (minTime == 0)
                        {
                            minTime = unixReceivedTime
                        }
                        else
                        {
                            if (unixReceivedTime < minTime)
                            {
                                minTime = unixReceivedTime
                            }
                        }
                    }
                }
            }
            if (minTime != 0){
                deltaArray.push(minTime - unixSentTime)
            }
        }
        var sum = 0
        for (let i = 0; i < deltaArray.length; i++)
        {
            sum += deltaArray[i]
        }
        if (deltaArray.length !== 0)
        {
            average = sum / deltaArray.length
            average = Math.trunc(average)
        }
        else
        {
            average = 0
        }
       
        top[key]['average'] = average
    }   
    return top
}

function getTopTen(recipients)
{
    var topRecipients = {}

    k = 0
    for (let key in recipients)
    {
       topRecipients[key] = recipients[key] 
       k++
       if (k == 10)
       {
        break
       }
    }
    return topRecipients
}

function getUnanswered(sentMails, receivedMails)
{
    ids = []
    sum = 0
    for (let i = 0; i<sentMails.length; i++)
    {   
        if (ids.includes(sentMails[i]['conversationId']))
        {
            continue
        }
        else
        {
            ids.push(sentMails[i]['conversationId'])
            conversationId = sentMails[i]['conversationId']
        }
        // Фильтруем отправленные сообщения по conversationId
        const sentMessages = sentMails.filter(mail => mail.conversationId === conversationId);
      
        // Фильтруем полученные сообщения по conversationId
        const receivedMessages = receivedMails.filter(mail => mail.conversationId === conversationId);
      
        // Считаем количество отправленных и полученных сообщений
        const sentCount = sentMessages.length;
        const receivedCount = receivedMessages.length;
      
        // Разница между отправленными и полученными сообщениями - неотвеченные сообщения
        const unansweredCount = sentCount - receivedCount;
        if (unansweredCount > 0){
            sum += unansweredCount  
        }
    }
    return sum
}

function getDomainFromEmail(email) {
    const parts = email.split('@');
    if (parts.length === 2) {
      return parts[1];
    }
    return null;
}

function getTopDomain(top)
{
    topDomains = {}
    for (let key in top)
    {
        domain = getDomainFromEmail(key)
        console.log(domain)
        if (domain in topDomains)
        {
            topDomains[domain] += top[key]['sent']
        }
        else
        {
            topDomains[domain] = top[key]['sent']
        }
    }
    const entries = Object.entries(topDomains);

    entries.sort((a, b) => b[1] - a[1]);

    topDomains = Object.fromEntries(entries);

    topDomains = getTopTen(topDomains)
    return topDomains
}

function saveDataToJson(data) {
    const json = JSON.stringify(data);
    console.log(json);

    writeFile("output.json", json, 'utf8', err => {
        if (err) {
            console.log("An error occured while writing JSON Object to File.");
            return console.log(err);
        }

        console.log("JSON file has been saved.");
    });
}

app.use(express.json());
app.use(express.static(__dirname));
app.get('/', (req, res) => {
//   const authCode = req.query.code; // Получение кода авторизации из параметров запроса

  // Ваша логика обработки кода авторизации здесь
  // Например, вы можете обменять код авторизации на токен доступа через MSAL и продолжить выполнение операций с API Graph

  // Возвращаем успешный ответ
  const filePath = path.join(__dirname, 'mstest.html');
  res.sendFile(filePath);
});

app.post('/metrics', (req, res) =>{
    const accessToken = req.body['token']
    const startDate = req.body['startDate']
    const endDate = req.body['endDate']
    var responseJson = {}
    getAllInfo(accessToken, startDate, endDate).then(result => {
        responseJson.sentMails = getSentMails(result['sentMails'])
        responseJson.receivedMails = getReceivedMails(result['receivedMails'])
        responseJson.averageReplyToSent = getReplyToSent(result['sentMails'], result['receivedMails'], 'sentDateTime', 'receivedDateTime')
        responseJson.averageReplyToReceived = getReplyToSent(result['receivedMails'], result['sentMails'], 'receivedDateTime', 'sentDateTime')
        responseJson.mailDayAnalytics = getDayMailAnalytics(result['sentMails'], 'sentDateTime')
        responseJson.receivedDayAnalytics = getDayMailAnalytics(result['receivedMails'], 'receivedDateTime')
        top = getTopRecipients(result['sentMails'], result['receivedMails'])
        top_Senders = getTopSenders(result['receivedMails'], top)

        responseJson.sentSum = getUnanswered(result['sentMails'], result['receivedMails'])
        responseJson.receivedSum = getUnanswered(result['receivedMails'], result['sentMails'])
        responseJson.topDomains = getTopDomain(top)
        top = getTopTen(top)
        responseJson.topRecipients = getReplyToSentByTop(result['sentMails'], result['receivedMails'], 'sentDateTime', 'receivedDateTime', top)
        responseJson.topSenders = getReplyToReceiveByTop(result['receivedMails'], result['sentMails'], 'receivedDateTime', 'sentDateTime', top_Senders)
        
        

        saveDataToJson(responseJson);
        res.status(200).json(responseJson)
    })

    
})

app.listen(port, () => {
  console.log(`Сервер запущен на порту ${port}`);
});
