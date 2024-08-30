#!/usr/bin/env node

import fs from "fs";
import { JSDOM } from "jsdom";
import xlsx from "xlsx";
import yargs from "yargs";

const { argv } = yargs(process.argv);

const html = fs.readFileSync(argv.html, "utf8");
// const html = fs.readFileSync("./data.html", "utf8");

const dom = new JSDOM(html);
const document = dom.window.document;

const workbook = xlsx.utils.book_new();

for (let i = 0; i < 5; i++) {
  const table = document.querySelector(`table:nth-of-type(${i + 1})`);
  const sheetname = `Sheet${i + 1}`;

  if (table) {
    const worksheet = xlsx.utils.table_to_sheet(table, {
      // cellDates: true,
      // UTC: true,
      raw: true,
    });
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetname);
  }
  // console.log("added sheet");
}

// for (let i = 0; i < workbook.SheetNames.length; i++) {
//   const sheet = workbook.Sheets[workbook.SheetNames[i]];
//   const range = xlsx.utils.decode_range(sheet["!ref"]);
//   for (let row = range.s.r + 1; row <= range.e.r; ++row) {
//     for (let column = range.s.c; column <= range.e.c; ++column) {
//       const ref = xlsx.utils.encode_cell({ r: row, c: column });
//       const cell = sheet[ref];
//       if (cell.t == "d") cell.t = "s";
//     }
//   }
// }

// const data = JSON.stringify(xlsx.utils.sheet_to_json(sheet));

// fs.writeFileSync(argv.file, data);

xlsx.writeFile(workbook, argv.spreadsheet);
// xlsx.writeFile(workbook, "./output.xlsx");
