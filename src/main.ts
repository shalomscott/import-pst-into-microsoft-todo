import 'isomorphic-fetch';
import { Client } from '@microsoft/microsoft-graph-client';
import DeviceCodeAuthProvider from './DeviceCodeAuthProvider';
import { PSTFile } from 'pst-extractor';
import { extractTasks, toDateTime, toRecurrence } from './helpers';

const { PST_FILE_LOCATION, APP_CLIENT_ID } = process.env;

(async function () {
	try {
		const client = Client.initWithMiddleware({
			authProvider: new DeviceCodeAuthProvider(APP_CLIENT_ID),
			defaultVersion: 'beta',
		});
		const userDetails = await client.api('/me').get();
		console.log(`Logged into account of: ${userDetails.displayName}`);

		const pstFile = new PSTFile(PST_FILE_LOCATION);
		const pstTasks = extractTasks(pstFile).filter(
			(task) => !task.isTaskComplete
		);
		console.log(
			`Extracted ${pstTasks.length} uncompleted tasks from PST file`
		);

		const taskList = await client
			.api('/me/todo/lists')
			.post({ displayName: 'PST Imported' });
		console.log(`Created new todo list: ${taskList.displayName}`);

		for (const pstTask of pstTasks) {
			const taskBody = {
				body: {
					content: pstTask.body,
					contentType: 'text',
				},
				dueDateTime: toDateTime(pstTask.taskDueDate),
				importance: ['low', 'normal', 'high'][pstTask.importance],
				isReminderOn: pstTask.reminderSet,
				recurrence: toRecurrence(pstTask.taskRecurrencePattern),
				reminderDateTime: toDateTime(
					new Date(
						pstTask.taskStartDate.getTime() -
							pstTask.reminderDelta * 60 * 1000
					)
				),
				title: pstTask.subject,
				createdDateTime: pstTask.creationTime.toISOString(),
				lastModifiedDateTime: pstTask.modificationTime.toISOString(),
			};
			console.log('Task body is:', JSON.stringify(taskBody, null, 2));
			const createdTask = await client
				.api(`/me/todo/lists/${taskList.id}/tasks`)
				.post(taskBody);
			console.log('Successfully created task: ' + createdTask.title);
		}
	} catch (e) {
		console.error(e);
	}
})();
