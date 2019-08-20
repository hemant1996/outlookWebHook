const express = require('express');
const router = express.Router();
const authHelper = require('../helpers/auth');

router.get('/', async function(req, res, next) {
  let parms = { title: 'Home', active: { home: true } };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;
    parms.debug = `User: ${userName}\nAccess Token: ${accessToken}`;
  } else {
    parms.signInUrl = authHelper.getAuthUrl();
    parms.debug = parms.signInUrl;
  }

  res.render('index', parms);
});

module.exports = router;