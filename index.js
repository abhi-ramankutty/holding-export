const express = require("express");
const exceljs = require("exceljs");
const fs = require("fs");
const _ = require('underscore');

const app = express();
const PORT = 4200;

app.get("/export", async (request, result) => {
    const workbook = new exceljs.Workbook();
    const sheet = workbook.addWorksheet("Holdings");

    sheet.columns = [
        { header: "Name", key: "name" },
        { header: "Category", key: "category" },
        { header: "Value", key: "value" }
    ];

    // const dataObj = JSON.parse(fs.readFileSync("holdings.json", "utf-8"));
    // await dataObj.result.map((item) => {
    //     const name = '';
    //     const category = '';
    //     const value = '';

    //     sheet.addRow({
    //         name,
    //         category,
    //         value
    //     })

    //     result.setHeader(
    //         "Content-type",
    //         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    //     );
    //     result.setHeader(
    //         "Content-disposition",
    //         "attachment; filename=TickertapeMFScreener.xlsx"
    //     );

    //     workbook.xlsx.write(result);
    // })

    const dataObj = JSON.parse(await fs.promises.readFile("holdings.json", "utf-8"));
    const group = _.groupBy(dataObj.holdings, 'tradingsymbol');
    _.each(group, (groupItem, itemName) => {
        const name = itemName;
        let category = groupItem[0].cap;
        let value = 0;

        _.each(groupItem, (item) => {
            value += (item.ltp * item.total_quantity)
        });

        if(category== 'L' || category== 'M' || category== 'S') {
            sheet.addRow({
                name,
                category,
                value
            })
        }

    });

    result.setHeader(
        "Content-type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    result.setHeader(
        "Content-disposition",
        "attachment; filename=Holdings.xlsx"
    );
    workbook.xlsx.write(result);

});

app.listen(PORT, () => {
    console.log("App is running on http://localhost:4200");
    console.log("Navigate to http://localhost:4200/export to download your file");
});
