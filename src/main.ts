import 'isomorphic-fetch';
import { Client } from '@microsoft/microsoft-graph-client';
// import extractTasks from './extract-tasks';
import AuthFromEnv from './auth-from-env';

console.log(process.env.MS_ACCESS_TOKEN);

(async function () {
	try {
		const client = Client.initWithMiddleware({
			authProvider: new AuthFromEnv(),
		});
		let userDetails = await client.api('/me').get();
		console.log(userDetails);
	} catch (e) {
		console.error(e);
	}
})();
