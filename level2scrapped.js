/////////////////// input the following steps to run the file//////////////
//npm install jsdom
//npm install excel4node
//npm install axios
///////////////////--------------x---------------------------///

let fs = require("fs");
let jsdom = require("jsdom");
let excel = require("excel4node");
let axios = require("axios");

// making a promise to download the data from pepcoding website.
let arrlinks = [
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/recursion-and-backtracking",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/bit-manipulation",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/dynamic-programming",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/graphs",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/hashmap-and-heaps",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/trees",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/linked-list",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/stacks",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/trie",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/arrays-and-strings",
  "https://www.pepcoding.com/resources/data-structures-and-algorithms-in-java-levelup/searching-and-sorting",
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

let sheet = workbook.addWorksheet("Level2");
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
      workbook.write("level2.xlsx");

      setTimeout(() => {}, 200);
    })
    .catch(function (err) {
      console.log(err);
      console.log("couldn't full fill the promise");
    });
}
