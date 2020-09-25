import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';

export default class AuthFromEnv implements AuthenticationProvider {
	public async getAccessToken(): Promise<string> {
		const { MS_ACCESS_TOKEN } = process.env;
		if (MS_ACCESS_TOKEN) {
			return MS_ACCESS_TOKEN;
		}
		throw Error('Missing MS_ACCESS_TOKEN env variable.');
	}
}
