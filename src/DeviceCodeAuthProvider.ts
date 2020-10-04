import { promises as fs } from 'fs';
import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import { PublicClientApplication } from '@azure/msal-node';

const ACCESS_TOKEN_PATH = './.access_token';

export default class DeviceCodeAuthProvider implements AuthenticationProvider {
	private accessToken: string;

	constructor(private clientId: string) {}

	public async getAccessToken(): Promise<string> {
		if (!this.accessToken) {
			try {
				this.accessToken = await fs.readFile(
					ACCESS_TOKEN_PATH,
					'utf-8'
				);
				console.log(
					'Using cached access token from ' + ACCESS_TOKEN_PATH
				);
			} catch {
				const client = new PublicClientApplication({
					auth: {
						clientId: this.clientId,
					},
				});

				const deviceCodeRequest = {
					deviceCodeCallback: (response) =>
						console.log(response.message),
					scopes: ['User.Read', 'Tasks.ReadWrite'],
				};

				const response = await client.acquireTokenByDeviceCode(
					deviceCodeRequest
				);

				this.accessToken = response.accessToken;
				await fs.writeFile(
					ACCESS_TOKEN_PATH,
					this.accessToken,
					'utf-8'
				);
			}
		}
		return this.accessToken;
	}
}
