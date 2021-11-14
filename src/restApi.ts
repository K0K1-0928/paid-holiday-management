const doGet = (e) => {
  let name = e.parameter.name;
  let records = getRecords();
  let record = records
    .filter((record) => record[eMailKey] === `${name}@evolveit.jp`)
    .shift();
  return ContentService.createTextOutput(JSON.stringify(record)).setMimeType(
    ContentService.MimeType.JSON
  );
};

const doPost = (e) => doGet(e);
