"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.BuilderXcell = void 0;
const adm_zip_1 = __importDefault(require("adm-zip"));
const fast_xml_parser_1 = require("fast-xml-parser");
const optionsParse = {
    ignoreAttributes: false,
};
const optionsBuild = {
    ignoreAttributes: false,
};
class BuilderXcell {
    constructor(path_template_xlsx = "./template/empty.xlsx") {
        this.si = [];
        const fastXML = new fast_xml_parser_1.XMLParser(optionsBuild);
        const empty_xlsx = new adm_zip_1.default(path_template_xlsx);
        const zipEntries = empty_xlsx.getEntries();
        this.parseToObj = zipEntries.map((zipEntry) => {
            const read = zipEntry.getData().toString("utf8");
            const a = fastXML.parse(read);
            const obj = {
                name: zipEntry.entryName,
                data: a,
            };
            return obj;
        });
    }
    saveFile(path_save_xlsx) {
        const new_xlsx = new adm_zip_1.default();
        const builder = new fast_xml_parser_1.XMLBuilder(optionsBuild);
        this.parseToObj.push({
            name: "xl/sharedStrings.xml",
            data: this.getSharedString(),
        });
        this.parseToObj.forEach((e) => {
            console.log(e.name);
            new_xlsx.addFile(e.name, Buffer.from(builder.build(e.data)), "");
        });
        new_xlsx.writeZip(path_save_xlsx);
    }
    getSharedString() {
        return {
            "?xml": {
                "@_version": "1.0",
                "@_encoding": "UTF-8",
                "@_standalone": "yes",
            },
            sst: {
                si: this.si,
                "@_xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "@_count": String(this.si.length),
                "@_uniqueCount": String(this.si.length),
            },
        };
    }
    addSharedString(text) {
        const space = typeof text === "string" && text.includes(" ");
        if (space) {
            const find = this.si.findIndex((s) => typeof s.t === "object" && s.t["#text"] === text);
            if (find >= 0) {
                return find;
            }
            else {
                this.si.push({ t: { "@_xml:space": "preserve", "#text": text } });
                return this.si.length - 1;
            }
        }
        else {
            const find = this.si.findIndex((s) => s.t === text);
            if (find >= 0) {
                return find;
            }
            else {
                this.si.push({ t: text });
                return this.si.length - 1;
            }
        }
    }
    addCell(row, col, text) {
        const sheet = this.parseToObj.find((e) => e.name === "xl/worksheets/sheet1.xml");
        if (sheet) {
            console.log(sheet.data.worksheet);
            if (!sheet.data.worksheet.sheetData) {
                sheet.data.worksheet.sheetData = { row: [] };
            }
            const frow = sheet.data.worksheet.sheetData.row.find((e) => e["@_r"] === String(row + 1));
            if (!frow) {
                sheet.data.worksheet.sheetData.row.push({
                    "@_r": String(row + 1),
                    c: [
                        {
                            "@_r": String.fromCharCode(65 + col) + String(row + 1),
                            "@_t": "s",
                            v: this.addSharedString(text),
                        },
                    ],
                });
            }
            console.log(sheet.data.worksheet.sheetData);
            sheet.data.worksheet.sheetData.row[row].c.push({
                "@_r": String.fromCharCode(65 + col) + String(row + 1),
                "@_t": "s",
                v: this.addSharedString(text),
            });
        }
    }
}
exports.BuilderXcell = BuilderXcell;
