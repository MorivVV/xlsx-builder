// import {
//   X2jOptions,
//   XMLBuilder,
//   XmlBuilderOptions,
//   XMLParser,
// } from "fast-xml-parser";
// import AdmZip from "adm-zip";
import { BuilderXcell } from "./classes/BuilderXcell";
// var empty_xlsx = new AdmZip("./template/cron_script.xlsx");
// var zipEntries = empty_xlsx.getEntries();

// const optionsParse: X2jOptions = {
//   ignoreAttributes: false,
// };
// const optionsBuild: XmlBuilderOptions = {
//   ignoreAttributes: false,
// };
// const fastXML = new XMLParser(optionsBuild);
// const builder = new XMLBuilder(optionsBuild);
// const new_xlsx = new AdmZip();
// const sharedString = {
//   "?xml": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" },
//   sst: {
//     si: [{ t: "строка которая пойдет в shARED" }],
//     "@_xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
//     "@_count": "1",
//     "@_uniqueCount": "1",
//   },
// };
// const parseToObj = zipEntries.map((zipEntry) => {
//   console.log(zipEntry.entryName); // outputs zip entries information
//   const read = zipEntry.getData().toString("utf8");
//   const a = fastXML.parse(read);
//   const obj = {
//     name: zipEntry.entryName,
//     data: a,
//   };

//   console.log(JSON.stringify(a));
//   console.log(builder.build(a));
//   //   new_xlsx.addFile(zipEntry.entryName, Buffer.from(builder.build(a)), "");
//   return obj;
// });

// parseToObj.forEach((e) => {
//   console.log(e.name);

//   new_xlsx.addFile(e.name, Buffer.from(builder.build(e.data)), "");
// });
// new_xlsx.writeZip("test.xlsx");

const ax = new BuilderXcell();
ax.addCell(0, 1, "Привет");
ax.addCell(0, 2, "Как дела");
ax.addCell(0, 3, "Ну и что");
ax.addCell(1, 2, "Ну и что22");
ax.addCell(1, 3, "Ну и что22");
ax.addCell(1, 4, "Как дела");
ax.saveFile("test34.xlsx");
