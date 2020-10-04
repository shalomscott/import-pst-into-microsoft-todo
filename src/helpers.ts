import { PSTFile, PSTFolder, PSTTask } from 'pst-extractor';
import {
	RecurrencePattern,
	WeekSpecific,
	MonthNthSpecific,
	PatternType,
	NthOccurrence,
	EndType,
	RecurFrequency,
} from 'pst-extractor/dist/RecurrencePattern.class';

const weekdays = [
	'sunday',
	'monday',
	'tuesday',
	'wednesday',
	'thursday',
	'friday',
	'saturday',
];

export function extractTasks(pstFile: PSTFile): PSTTask[] {
	return getTasks(pstFile.getRootFolder());
}

function getTasks(folder: PSTFolder): PSTTask[] {
	const tasks: PSTTask[] = [];

	if (folder.hasSubfolders) {
		for (const subFolder of folder.getSubFolders()) {
			tasks.push(...getTasks(subFolder));
		}
	}

	let task = folder.getNextChild();
	while (task !== null) {
		if (task.messageClass === 'IPM.Task') {
			tasks.push(task);
		}
		task = folder.getNextChild();
	}

	return tasks;
}

export function toDateTime(date: Date) {
	return {
		dateTime: date.toISOString().split('.')[0],
		timeZone: 'Etc/GMT',
	};
}

export function toRecurrence(recurrencePattern: RecurrencePattern) {
	if (!recurrencePattern) return undefined;

	const pattern: any = {};
	const range: any = {};

	switch (recurrencePattern.patternType) {
		case PatternType.Day:
			pattern.type = 'daily';
			pattern.interval = recurrencePattern.period / (24 * 60);
			break;
		case PatternType.Week:
			pattern.type = 'weekly';
			pattern.interval = recurrencePattern.period;
			pattern.daysOfWeek = (recurrencePattern.patternTypeSpecific as WeekSpecific)
				.map((day, i) => (day ? weekdays[i] : null))
				.filter((day) => !!day);
			pattern.firstDayOfWeek = weekdays[recurrencePattern.firstDOW];
			break;
		case PatternType.Month:
			switch (recurrencePattern.recurFrequency) {
				case RecurFrequency.Monthly:
					pattern.type = 'absoluteMonthly';
					pattern.interval = recurrencePattern.period;
					break;
				case RecurFrequency.Yearly:
					pattern.type = 'absoluteYearly';
					pattern.interval = recurrencePattern.period / 12;
					pattern.month = recurrencePattern.startDate.getMonth() + 1;
					break;
			}
			pattern.dayOfMonth = recurrencePattern.patternTypeSpecific as number;
			break;
		case PatternType.MonthNth:
			switch (recurrencePattern.recurFrequency) {
				case RecurFrequency.Monthly:
					pattern.type = 'relativeMonthly';
					pattern.interval = recurrencePattern.period;
					break;
				case RecurFrequency.Yearly:
					pattern.type = 'relativeYearly';
					pattern.interval = recurrencePattern.period / 12;
					pattern.month = recurrencePattern.startDate.getMonth() + 1;
					break;
			}
			pattern.daysOfWeek = (recurrencePattern.patternTypeSpecific as MonthNthSpecific).weekdays
				.map((day, i) => (day ? weekdays[i] : null))
				.filter((day) => !!day);
			pattern.index = NthOccurrence[
				(recurrencePattern.patternTypeSpecific as MonthNthSpecific).nth
			].toLowerCase();
			break;
	}

	range.startDate = recurrencePattern.startDate.toISOString().split('T')[0];
	switch (recurrencePattern.endType) {
		case EndType.AfterDate:
			range.type = 'endDate';
			range.endDate = recurrencePattern.endDate
				.toISOString()
				.split('T')[0];
			break;
		case EndType.AfterNOccurrences:
			range.type = 'numbered';
			range.numberOfOccurrences = recurrencePattern.occurrenceCount;
			break;
		case EndType.NeverEnd:
			range.type = 'noEnd';
			break;
	}

	return { pattern, range };
}
