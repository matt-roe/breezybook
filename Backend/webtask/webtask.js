'use latest';

import express from 'express';
import { fromExpress } from 'webtask-tools';
import bodyParser from 'body-parser';
import stripe from 'stripe';

bodyParser.urlencoded();

var app = express();
app.use(bodyParser.urlencoded());

app.post('/payment', (req,res) => {
  var ctx = req.webtaskContext;
  var STRIPE_SECRET_KEY = ctx.secrets.STRIPE_SECRET_KEY;

  stripe(STRIPE_SECRET_KEY).charges.create({
    amount: req.query.amount,
    currency: req.query.currency,
    source: req.body.stripeToken,
    description: req.query.description
  },
  payment_intent_data: {
    application_fee_amount: 100
  },
   (err, charge) => {
    const status = err ? 400: 200;
    const message = err ? err.message: 'Payment done!';
    res.writeHead(status, { 'Content-Type': 'text/html' });
    return res.end('<h1>' + message + '</h1>');
  },
  {
  stripe_account: CONNECTED_STRIPE_ACCOUNT_ID,
  });
});

module.exports = fromExpress(app);
