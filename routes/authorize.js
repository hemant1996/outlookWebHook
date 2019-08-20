const express = require('express');
const router = express.Router();
const authHelper = require('../helpers/auth');

router.get('/', async function (req, res, next) {
    const code = req.query.code;
    if (code) {
        try {
            await authHelper.getTokenFromCode(code, res);
            res.redirect('/');
        } catch (error) {
            res.render('error', { title: 'Error', message: 'Error exchanging code for token', error: error });
        }
    } else {
        res.render('error', { title: 'Error', message: 'Authorization error', error: { status: 'Missing code parameter' } });
    }
});

router.get('/signout', function (req, res, next) {
    authHelper.clearCookies(res);

    res.redirect('/');
});

module.exports = router;