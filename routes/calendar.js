var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');
var moment = require('moment')
var listenRouter = express.Router();

router.get('/', async function (req, res, next) {
    let parms = { title: 'Calendar', active: { calendar: true } };

    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    const userName = req.cookies.graph_user_name;

    if (accessToken && userName) {
        parms.user = userName;

        const client = graph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        const start = new Date(new Date().setHours(0, 0, 0));
        const end = new Date(new Date(start).setDate(start.getDate() + 7));

        try {
            const result = await client
                .api(`/me/calendarView?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}`)
                .top(10)
                .select('subject,start,end,attendees')
                .orderby('start/dateTime DESC')
                .get();

            for (let i = 0; i < result.value.length; i++) {
                result.value[i].start.dateTime = moment(result.value[i].start.dateTime).format("YYYY-MM-DDTkk:mm")
                result.value[i].end.dateTime = moment(result.value[i].end.dateTime).format("YYYY-MM-DDTkk:mm")
            }
            parms.events = result.value;
            res.render('calendar', parms);
        } catch (err) {
            parms.message = 'Error retrieving events';
            parms.error = { status: `${err.code}: ${err.message}` };
            parms.debug = JSON.stringify(err.body, null, 2);
            res.render('error', parms);
        }

    } else {
        res.redirect('/');
    }
});

router.post('/update', async function (req, res, next) {
    let parms = { title: 'Calendar', active: { calendar: true } };

    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    const userName = req.cookies.graph_user_name;

    if (accessToken && userName) {
        parms.user = userName;

        const client = graph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        let id = req.body.id
        try {
            const result = await client
                .api(`/me/events/${id}`)
                .patch({
                    subject: req.body.subject
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

router.get('/subscribe', async function (req, res, next) {
    let parms = { title: 'Calendar', active: { calendar: true } };

    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    const userName = req.cookies.graph_user_name;

    if (accessToken && userName) {
        parms.user = userName;

        const client = graph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        let id = req.body.id
        try {
            const result = await client
                .api(`https://graph.microsoft.com/v1.0/subscriptions`)
                .post({
                    changeType: "created,updated",
                    notificationUrl: "https://5cf7a123.ngrok.io/calendar/notificationClient",
                    resource: "/me/mailfolders('inbox')/messages",
                    expirationDateTime: new Date(moment().add('H',2)).toISOString(),
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
    if(req.query && req.query.validationToken){
        res.set('Content-Type', 'plain/text');
        res.send(req.query.validationToken);
    } else {
        res.status(500).send('invalid token')
    }
});
module.exports = router;