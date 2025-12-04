#!/usr/bin/env ts-node
import { PSTFile } from 'pst-extractor';
import { RecurrencePattern } from 'pst-extractor/dist/RecurrencePattern.class';

async function checkEvent() {
  try {
    const pstFile = new PSTFile('prod/C-Users-eshc3-backup.pst');

    console.log('Opening PST file...\n');

    // Find calendar folder
    let calendarFolder: any = null;
    const rootFolder = pstFile.getRootFolder();

    function findAllCalendars(folder: any, calendars: any[] = []): any[] {
      const containerClass = folder.containerClass?.trim().toLowerCase();
      if (containerClass?.startsWith('ipf.appointment')) {
        calendars.push(folder);
      }
      if (folder.hasSubfolders) {
        const subFolders = folder.getSubFolders();
        for (const subFolder of subFolders) {
          findAllCalendars(subFolder, calendars);
        }
      }
      return calendars;
    }

    const calendarFolders = findAllCalendars(rootFolder);

    if (calendarFolders.length === 0) {
      console.log('No calendar folders found!');
      return;
    }

    console.log('Found ' + calendarFolders.length + ' calendar folder(s)\n');

    // Search for the event in all calendar folders
    let found = false;

    for (const folder of calendarFolders) {
      console.log('Searching in: ' + folder.displayName + ' (' + folder.contentCount + ' items)');

      let item = folder.getNextChild();
      while (item) {
        if (item.subject && item.subject.includes('Photo Course')) {
        console.log('=== FOUND: ' + item.subject + ' ===\n');
        console.log('Subject: ' + item.subject);
        console.log('Start Time: ' + item.startTime);
        console.log('End Time: ' + item.endTime);

        if (item.startTime && item.endTime) {
          const duration = (item.endTime.getTime() - item.startTime.getTime()) / (1000 * 60 * 60 * 24);
          console.log('Duration: ' + duration.toFixed(2) + ' days');
        }

        console.log('Duration (minutes): ' + item.duration);
        console.log('Is Recurring: ' + item.isRecurring);
        console.log('All Day Event: ' + item.allDayEvent);
        console.log('Location: ' + (item.location || 'N/A'));

        if (item.recurrenceStructure && Buffer.isBuffer(item.recurrenceStructure)) {
          console.log('\n--- Recurrence Pattern Details ---');
          const pattern = new RecurrencePattern(item.recurrenceStructure);
          const patternJSON = pattern.toJSON();

          console.log('Recurrence Type: ' + pattern.recurFrequency + ' (8202=DAILY, 8203=WEEKLY, 8204=MONTHLY, 8205=YEARLY)');
          console.log('Pattern Type: ' + pattern.patternType);
          console.log('Period: ' + pattern.period);
          console.log('End Type: ' + pattern.endType + ' (8225=AfterDate, 8226=AfterCount, 8227=NeverEnd)');
          console.log('End Date: ' + pattern.endDate);
          console.log('Occurrence Count: ' + pattern.occurrenceCount);
          console.log('First DateTime: ' + pattern.firstDateTime);
          console.log('\nFull Pattern JSON:');
          console.log(JSON.stringify(patternJSON, null, 2));
        } else {
          console.log('\nNo recurrence structure found');
        }

          found = true;
          break;
        }
        item = folder.getNextChild();
      }

      if (found) break;
    }

    if (!found) {
      console.log('\nEvent not found in any calendar folder!');
    }

    pstFile.close();

  } catch (error) {
    console.error('Error:', error);
  }
}

checkEvent();
