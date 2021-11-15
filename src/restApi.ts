const doGet = (e) => {
  let returnJson = { member: null };
  let records = getRecords();

  if (e && e.parameter && e.parameter.name) {
    let name = e.parameter.name;
    returnJson.member = records.filter(
      (record) => record[eMailKey] === `${name}@evolveit.jp`
    );
  } else {
    returnJson.member = records;
  }

  return ContentService.createTextOutput(
    JSON.stringify(returnJson)
  ).setMimeType(ContentService.MimeType.JSON);
};

const doPost = (e) => doGet(e);
