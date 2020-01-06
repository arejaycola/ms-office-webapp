var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

/* GET /mail */
router.get('/', async function(req, res, next) {
	let parms = { title: 'Inbox', active: { inbox: true } };

	const accessToken = await authHelper.getAccessToken(req.cookies, res);
	const userName = req.cookies.graph_user_name;

	if (accessToken && userName) {
		parms.user = userName;

		// Initialize Graph client
		const client = graph.Client.init({
			authProvider: done => {
				done(null, accessToken);
			}
		});

		try {
			// Get the 10 newest messages from inbox
			const result = await client
				.api('/me/mailfolders/inbox/messages')
				.top(10)
				.select('subject,from,receivedDateTime,isRead')
				.orderby('receivedDateTime DESC')
				.get();

			parms.messages = result.value;
			res.render('mail', parms);
		} catch (err) {
			parms.message = 'Error retrieving messages';
			parms.error = { status: `${err.code}: ${err.message}` };
			parms.debug = JSON.stringify(err.body, null, 2);
			res.render('error', parms);
		}
	} else {
		// Redirect to home
		res.redirect('/');
	}
});

router.post('/send', async function(req, res) {
	try {
		console.log('Sending...');
		const accessToken = await authHelper.getAccessToken(req.cookies, res);

		const client = graph.Client.init({
			authProvider: done => {
				done(null, accessToken);
			}
		});

		const sendMail = {
			message: {
				subject: 'Meet for lunch?',
				body: {
					contentType: 'Text',
					content: 'The new cafeteria is open.'
				},
				toRecipients: [
					{
						emailAddress: {
							address: 'gabe.lewis@cybervizgroup.onmicrosoft.com'
						}
					}
				]
			},
			saveToSentItems: 'true'
		};

		let response = await client.api('/me/sendMail').post(sendMail);
		console.log(response);
	} catch (e) {
		console.log(e);
	}
});

module.exports = router;
