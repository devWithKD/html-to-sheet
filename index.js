#!/usr/bin/env node

import fs from "fs";
import { JSDOM } from "jsdom";
import xlsx from "xlsx";
import yargs from "yargs";

const { argv } = yargs(process.argv);

const html = fs.readFileSync(argv.html, "utf8");

const dom = new JSDOM(html);
const document = dom.window.document;

const workbook = xlsx.utils.book_new();

for (let i = 0; i < 5; i++) {
  const table = document.querySelector(`table:nth-of-type(${i + 1})`);
  const sheetname = `Sheet${i + 1}`;

  if (table) {
    const worksheet = xlsx.utils.table_to_sheet(table, {
      cellDates: true,
      // UTC: true,
      // raw: true,
    });
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetname);
  }
  // console.log("added sheet");
}

xlsx.writeFile(workbook, argv.spreadsheet);
