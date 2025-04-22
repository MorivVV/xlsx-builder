"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const fast_xml_parser_1 = require("fast-xml-parser");
const adm_zip_1 = __importDefault(require("adm-zip"));
var empty_xlsx = new adm_zip_1.default("./template/empty.xlsx");
var zipEntries = empty_xlsx.getEntries();
const optionsParse = {
    ignoreAttributes: false,
};
const optionsBuild = {
    ignoreAttributes: false,
};
const fastXML = new fast_xml_parser_1.XMLParser(optionsBuild);
const builder = new fast_xml_parser_1.XMLBuilder(optionsBuild);
const new_xlsx = new adm_zip_1.default();
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
