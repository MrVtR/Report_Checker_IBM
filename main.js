const fs = require('fs');
const XLSX = require('xlsx');
const pressAnyKey = require('press-any-key');
const prompt = require('prompt-sync')({ sigint: true });
const figlet = require('figlet');
const multiLinePrompt = (ask) => {
  const lines = ask.split(/\r?\n/);
  const promptLine = lines.pop();
  console.log(lines.join('\n'));
  return prompt(promptLine);
};

console.log('Script made by:');
figlet.text(
  'Vítor R. G. Gomes - IBM',
  {
    horizontalLayout: 'fitted',
    verticalLayout: 'fitted',
    width: 80,
    whitespaceBreak: true,
  },
  function (err, data) {
    console.log(data);
    callListOption = multiLinePrompt(
      `Which Call List do you want to check?\n1- BLD\n2- Dev\n3- HLI\n4- HFD\nType one option: `,
    );
    let callListPath = '';
    switch (callListOption) {
      case '1':
        callListPath = 'files/callList1.xlsx';
        break;
      case '2':
        callListPath = 'files/callList2.xlsx';
        break;
      case '3':
        callListPath = 'files/callList3.xlsx';
        break;
      case '4':
        callListPath = 'files/callList4.xlsx';
        break;
      default:
        console.log('\nInvalid option, terminating the program...');
        process.exit(1);
    }
    pressAnyKey(
      '\nPut the file report.txt, referring to the Call List that will be checked, in the path: "/files/report.txt".\nAfter that, press any key to continue the program or CTRL+C to exit\n',
      {
        ctrlC: 'reject',
      },
    )
      .then(() => {
        console.log('\nStarting script...');
        //Getting data from report.txt
        fs.readFile('files/report.txt', function (err, data) {
          if (err) throw err;
          const oldArr = data.toString().replace(/\r\n/g, '\n').split('\n');
          const reportArr = trimSortArr(oldArr);

          //Getting data from callList xlsx
          var workbook = XLSX.readFile(callListPath);
          const sheetName = workbook.SheetNames[0];
          const workbookSheet = workbook.Sheets[sheetName];
          let sheetRange = XLSX.utils.decode_range(workbookSheet['!ref']);
          let callListArr = [];
          for (let i = 1; i <= sheetRange.e.r + 1; i++) {
            callListArr.push(workbookSheet[`A${i}`].v);
          }
          callListArr = trimSortArr(callListArr);

          let CheckSheetToReport = [];
          let CheckReportToSheet = [];

          console.log('\nSearching...');
          for (let i = 0; i < callListArr.length; i++)
            if (!reportArr.includes(callListArr[i]))
              CheckSheetToReport.push(callListArr[i]);
          for (let i = 0; i < reportArr.length; i++)
            if (!callListArr.includes(reportArr[i]))
              CheckReportToSheet.push(reportArr[i]);

          if (CheckReportToSheet.length == 0)
            CheckReportToSheet.push(
              'All jobs contained in the Report are in the Call List',
            );
          if (CheckSheetToReport.length == 0)
            CheckSheetToReport.push(
              'All jobs contained in the Call List are in the Report',
            );

          console.log('\nSearch completed, generating results file...');
          writeFile('ResultsSheetToReport', CheckSheetToReport);
          writeFile('ResultsReportToSheet', CheckReportToSheet);
          console.log('\nSee ya next time! :D');
        });
      })
      .catch(() => {
        console.log('\nExiting the program, see you later...');
      });
  },
);

function trimSortArr(oldArr) {
  let trimmedArr = oldArr.map((str) => str.trim());
  const ne = (str) => str.replace(/\d+/g, (n) => n.padStart(8, '0'));
  const nc = (a, b) => ne(a).localeCompare(ne(b));
  trimmedArr.sort(nc);
  if (trimmedArr[0] == '') trimmedArr.shift();
  return trimmedArr;
}
function writeFile(name, arr) {
  var file = fs.createWriteStream(name + '.txt');
  file.on('error', function (err) {
    console.log(err);
  });
  arr.forEach(function (v) {
    file.write(v + '\n');
  });
  console.log('\nFile: ' + name + '.txt successfully created');
  file.end();
}
