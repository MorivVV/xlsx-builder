import {
  X2jOptions,
  XMLBuilder,
  XmlBuilderOptions,
  XMLParser,
} from "fast-xml-parser";
import AdmZip from "adm-zip";
var empty_xlsx = new AdmZip("./template/empty.xlsx");
var zipEntries = empty_xlsx.getEntries();

const optionsParse: X2jOptions = {
  ignoreAttributes: false,
};
const optionsBuild: XmlBuilderOptions = {
  ignoreAttributes: false,
};
const fastXML = new XMLParser(optionsBuild);
const builder = new XMLBuilder(optionsBuild);
const new_xlsx = new AdmZip();
zipEntries.forEach(function (zipEntry) {
  console.log(zipEntry.entryName); // outputs zip entries information

  console.log(zipEntry.getData().toString("utf8"));
  const a = fastXML.parse(zipEntry.getData().toString("utf8"));
  if (zipEntry.entryName.includes("xl/worksheets/sheet1.xml")) {
    a.worksheet.sheetData = {
      row: {
        "@_r": "1",
        "@_spans": "1:1",
        "@_x14ac:dyDescent": "0.25",

        c: { "@_r": "A1", v: 1234 },
      },
    };
  }
  console.log(a);
  console.log(builder.build(a));
  new_xlsx.addFile(zipEntry.entryName, Buffer.from(builder.build(a)), "");
});
new_xlsx.writeZip("test.xlsx");
