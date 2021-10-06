/////////////////// input the following steps to run the file//////////////
//npm install jsdom
//npm install excel4node
//npm install axios
//node index.js
///////////////////--------------x---------------------------///

let fs = require("fs");
let jsdom = require("jsdom");
let excel = require("excel4node");
let axios = require("axios");
const { title } = require("process");

// making a promise to download the data from pepcoding website.
let arrlinks = [
  "https://www.pepcoding.com/resources/online-java-foundation/getting-started",
  "https://www.pepcoding.com/resources/online-java-foundation/patterns",
  "https://www.pepcoding.com/resources/online-java-foundation/function-and-arrays",
  "https://www.pepcoding.com/resources/online-java-foundation/2d-arrays",
  "https://www.pepcoding.com/resources/online-java-foundation/string,-string-builder-and-arraylist",
  "https://www.pepcoding.com/resources/online-java-foundation/introduction-to-recursion",
  "https://www.pepcoding.com/resources/online-java-foundation/recursion-in-arrays",
  "https://www.pepcoding.com/resources/online-java-foundation/recursion-with-arraylists",
  "https://www.pepcoding.com/resources/online-java-foundation/recursion-on-the-way-up",
  "https://www.pepcoding.com/resources/online-java-foundation/recursion-backtracking",
  "https://www.pepcoding.com/resources/online-java-foundation/time-and-space-complexity",
  "https://www.pepcoding.com/resources/online-java-foundation/dynamic-programming-and-greedy",
  "https://www.pepcoding.com/resources/online-java-foundation/stacks-and-queues",
  "https://www.pepcoding.com/resources/online-java-foundation/linked-lists",
  "https://www.pepcoding.com/resources/online-java-foundation/generic-tree",
  "https://www.pepcoding.com/resources/online-java-foundation/binary-tree",
  "https://www.pepcoding.com/resources/online-java-foundation/binary-search-tree",
  "https://www.pepcoding.com/resources/online-java-foundation/hashmap-and-heap",
  "https://www.pepcoding.com/resources/online-java-foundation/graphs",
];

// excell sheet properties intialising

let workbook = new excel.Workbook();
let headStyle = workbook.createStyle({
  font: {
    bold: true,
    size: 16,
  },
});

let ContentStylequestion = workbook.createStyle({
  font: {
    size: 14,
    color: "blue",
  },
});

let ContentStyle = workbook.createStyle({
  font: {
    size: 14,
  },
});

let sheet = workbook.addWorksheet("Level1");
let sr = 2;
let sc = 2;
sheet.cell(2, 4).string(" Status[Yes/No] ").style(headStyle);

// getting data from web and putiing in excell sheet
for (let c = 0; c < arrlinks.length; c++) {
  let downloadHtmlPromise = axios.get(arrlinks[c]);
  downloadHtmlPromise
    .then(function (html) {
      let JSDOM = new jsdom.JSDOM(html.data);
      let document = JSDOM.window.document;

      // // storing all the content
      //   let title = document.querySelectorAll("title");
      // console.log(document.title);
      let topictitle = document.title;
      let title = topictitle.split("|");

      let StoreTheContent = document.querySelectorAll(
        "li[resource-Type=ojquestion] > div > a"
      );

      //   console.log(StoreTheContent.length);
      //   console.log(StoreTheContent[1].textContent.trim());
      //   console.log("https://www.pepcoding.com" + StoreTheContent[1].href);

      // writting data into sheet
      sheet.cell(sr, sc).string(title[1].trim()).style(headStyle);
      sr = sr + 2;
      for (let j = 0; j < StoreTheContent.length; j++) {
        sheet.cell(sr, 1).number(j + 1);
        sheet
          .cell(sr, 2)
          .link("https://www.pepcoding.com" + StoreTheContent[j].href)
          .string(StoreTheContent[j].textContent.trim())
          .style(ContentStylequestion);
        sheet.cell(sr, 4).string(" <-> ").style(ContentStyle);
        sr++;
      }
      sr = sr + 3;
      workbook.write("level1.xlsx");

      setTimeout(() => {}, 200);
    })
    .catch(function (err) {
      console.log(err);
      console.log("couldn't full fill the promise");
    });
}
