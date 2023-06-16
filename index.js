/*
* Title : Auto-reply to the mails which have no prior replies
* Author: Buddha Varsha
*/ 

// importing the required modules
const express = require('express')
const app = express()
const path = require('path')
const fs = require('fs').promises;
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');

// port number to listen
const port = 3000;


// required permissions to the API to access user's data
const SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://mail.google.com/'
];

// defining the root url for the application
app.get('/', async (req, res) => {

    //reading the client secrets from the local file through fs
    const credentials = await fs.readFile('credentials.json');

    //authorizing a client with credentials
    const auth = await authenticate({
        keyfilePath : path.join(__dirname, 'credentials.json'),
        scopes : SCOPES,
    });

    // creating the client object to interact with Gmail API
    const gmail = google.gmail({version: 'v1', auth});

    // to retrieve the list of labels 
    const response = await gmail.users.labels.list({
        userId: 'me',
    });

    //name of the label to be created
    const LABEL_NAME = 'Vacation';

    //load credentials form file
    async function loadCredentials(){
        const filePath = path.join(process.cwd(), 'credentials.json');
        const content = await fs.readFile(filePath, {encoding: 'utf-8'});
        return JSON.parse(content);
    }


    //retrieving and storing the messages which are unreplied in a list
    async function getUnrepliedMessages(auth){
        const gmail = google.gmail({version: 'v1', auth});
        const res = await gmail.users.messages.list({
            userId: 'me',
            labelIds: ['INBOX'],
            q: 'is:unread',
        });
        console.log(res.data.messages); // debug point to see the list of unreplied messages with their label id in the console
        return res.data.messages || []; // returns the list of unreplied messages, if there are no messages then it returns an empty array
    }


    //sending reply to the unread messages
    async function sendReply(auth, message){
        const gmail = google.gmail({version:'v1', auth});
        const res = await gmail.users.messages.get({
            userId: 'me',
            id: message.id,
            format: 'metadata',
            metadataHeaders: ['Subject', 'From'],
        });

        //retrieves the value of subject header
        const subject = res.data.payload.headers.find(
            (header) => header.name === 'Subject'
        ).value;

        //retrieves the value of from header
        const from = res.data.payload.headers.find(
            (header) => header.name === 'From'
        ).value;

        // initializing the variables
        const replyTo = from.match(/<(.*)>/)?.[1] || from;   //name <mail>
        const replySubject = subject.startsWith('Re:') ? subject : `Re: ${subject}`;
        const replyBody = `Hi, \n\n I'm currently on vacation and will get back to you soon. \n\nBest\nVarsha`;
        const rawMessage = [
            `From: me`,
            `To: ${replyTo}`,
            `Subject: ${replySubject}`,
            `In-Reply-To: ${message.id}`,
            `References: ${message.id}`,
            '',
            replyBody,
        ].join('\n');

        //raw message into base-64 encoded string
        const encodedMessage = Buffer.from(rawMessage).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
        await gmail.users.messages.send({
            userId: 'me',
            requestBody: {
                raw: encodedMessage,
            },
        });
    }

    
    //creating the label
    async function createLabel(auth){
        const gmail = google.gmail({version: 'v1', auth});
        try{
            const res = await gmail.users.labels.create({
                userId: 'me',
                requestBody: {
                    name: LABEL_NAME,
                    labelListVisibility: 'labelShow',
                    messageListVisibility: 'show',      
                },
            });
            return res.data.id;
        } catch (err){                        // exceptional handling 
            if(err.code === 409){
                const res = await gmail.users.labels.list({
                    userId: 'me',
                });
                const label = res.data.labels.find((label) => label.name === 'Vacation');
                return label.id;                         
            }
            else{
                throw err;
            }
        }
    }
    

    //adding label to the message and move it to the label folder
    async function addLabel(auth, message, labelId){
        const gmail = google.gmail({version: 'v1', auth});
        await gmail.users.messages.modify({
            userId: 'me',
            id: message.id,
            requestBody: {
                addLabelIds: [labelId],
                removeLabelIds: ['INBOX'],
            },
        });
    }

        //main function
        async function main(){

            //creating the label 
            const labelId = await createLabel(auth);
            console.log(`Created or found label with id ${labelId}`);


            //repeating the following steps in random intervals
            setInterval(async() => {

                //get unread messages
                const messages = await getUnrepliedMessages(auth);
                console.log(`Found ${messages.length} unreplied messages`);

                //for each message
                for (const message of messages){

                    //send reply to the message
                    await sendReply(auth, message);
                    console.log(`Sent reply to message with id ${message.id}`);

                    //add label to the message and move it to label folder
                    await addLabel(auth, message, labelId);
                    console.log(`Added label to message with id ${message.id}`);

                }
            }, Math.floor(Math.random() * (120 - 45 + 1) + 45) * 1000);  //random interval of 45 to 120 seconds as mentioned in the requirement
            
        }
    main().catch(console.error);
    

    const labels = response.data.labels;
    res.send('Successful');
});

app.listen(port, () => {
    console.log(`Listening at http://localhost:${port}`);
});
