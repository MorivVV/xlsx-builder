import AdmZip from "adm-zip";
import {
  X2jOptions,
  XMLBuilder,
  XmlBuilderOptions,
  XMLParser,
} from "fast-xml-parser";
const optionsParse: X2jOptions = {
  ignoreAttributes: false,
};
const optionsBuild: XmlBuilderOptions = {
  ignoreAttributes: false,
};
type TSist = {
  t:
    | string
    | number
    | boolean
    | { "#text": string; "@_xml:space"?: "preserve" };
};
export class BuilderXcell {
  parseToObj: {
    name: string;
    data: any;
  }[];
  si: TSist[] = [];

  constructor(path_template_xlsx: string = "./template/empty.xlsx") {
    const fastXML = new XMLParser(optionsBuild);
    const empty_xlsx: AdmZip = new AdmZip(path_template_xlsx);
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

  saveFile(path_save_xlsx: string) {
    const new_xlsx = new AdmZip();
    const builder = new XMLBuilder(optionsBuild);
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

  private addSharedString(text: string) {
    const space =
      typeof text === "string" && (text.startsWith(" ") || text.endsWith(" "));
    if (space) {
      const find = this.si.findIndex(
        (s) => typeof s.t === "object" && s.t["#text"] === text
      );
      if (find >= 0) {
        return find;
      } else {
        this.si.push({ t: { "@_xml:space": "preserve", "#text": text } });
        return this.si.length - 1;
      }
    } else {
      const find = this.si.findIndex((s) => s.t === text);
      if (find >= 0) {
        return find;
      } else {
        this.si.push({ t: text });
        return this.si.length - 1;
      }
    }
  }
  addSheet(name: string) {}
  getSheet(name: string = "xl/worksheets/sheet1.xml") {
    const sheet = this.parseToObj.find((e) => e.name === name);
    if (sheet) {
      if (!sheet.data.worksheet.sheetData) {
        sheet.data.worksheet.sheetData = { row: [] };
      }
      return sheet.data.worksheet;
    } else {
      return this.addSheet(name);
    }
  }
  addRow(row: number) {}
  getRow(row: number) {}
  addCell(row: number, col: number, text: string) {
    const sheet = this.getSheet("xl/worksheets/sheet1.xml");

    if (sheet) {
      console.log(sheet);

      const frow = sheet.sheetData.row.find(
        (e: { [x: string]: string }) => e["@_r"] === String(row + 1)
      );

      if (!frow) {
        sheet.sheetData.row.push({
          "@_r": String(row + 1),
          c: [
            {
              "@_r": String.fromCharCode(65 + col) + String(row + 1),
              "@_t": "s",
              v: this.addSharedString(text),
            },
          ],
        });
      } else {
        sheet.sheetData.row[row].c.push({
          "@_r": String.fromCharCode(65 + col) + String(row + 1),
          "@_t": "s",
          v: this.addSharedString(text),
        });
      }
      console.log(sheet.sheetData);
    }
  }
}
