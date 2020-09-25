import { PSTFile, PSTFolder, PSTTask } from 'pst-extractor';

export default function extractTasks(source: string): PSTTask[] {
	return getTasks(new PSTFile(source).getRootFolder());
}

function getTasks(folder: PSTFolder): PSTTask[] {
	const tasks: PSTTask[] = [];

	if (folder.hasSubfolders) {
		for (const subFolder of folder.getSubFolders()) {
			tasks.push(...getTasks(subFolder));
		}
	}

	let task: PSTTask = folder.getNextChild();
	while (task !== null) {
		if (task.messageClass === 'IPM.Task') {
			tasks.push(task);
		}
		task = folder.getNextChild();
	}

	return tasks;
}
