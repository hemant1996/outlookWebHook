var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');
var moment = require('moment')

router.get('/subscribe', async function (req, res, next) {
    let parms = { title: 'Calendar', active: { calendar: true } };
    let subscriptionType = req.query.type;
    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    const userName = req.cookies.graph_user_name;

    if (accessToken && userName) {
        parms.user = userName;

        const client = graph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
        let resource = "";
        if (subscriptionType == "mail") {
            resource = "/me/mailfolders('inbox')/messages";
            changeType = "created,updated";
        } else if (subscriptionType == "contacts") {
            resource = "me/contacts";
            changeType = "created,updated";
        } else if (subscriptionType == "events") {
            resource = "me/events";
            changeType = "created,updated";
        }
        try {
            await client
                .api(`https://graph.microsoft.com/v1.0/subscriptions`)
                .post({
                    changeType: changeType,
                    notificationUrl: "https://5cf7a123.ngrok.io/calendar/notificationClient",
                    resource: resource,
                    expirationDateTime: new Date(moment().add('H', 2)).toISOString(),
                    clientState: "SecretClientState"
                })
            res.redirect('/calendar/');
        } catch (err) {
            parms.message = 'Error retrieving events';
            parms.error = { status: `${err.code}: ${err.message}` };
            parms.debug = JSON.stringify(err.body, null, 2);
            res.render('error', parms);
        }

    } else {
        res.redirect('/calendar/');
    }
});

router.post('/notificationClient', (req, res, next) => {
    if (req.query && req.query.validationToken) {
        res.set('Content-Type', 'plain/text');
        res.send(req.query.validationToken);
    } else {
        res.status(500).send('invalid token')
    }
});
module.exports = router;