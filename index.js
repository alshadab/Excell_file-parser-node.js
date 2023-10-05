const xlsx = require("xlsx");

const excellParser = (model) => {
  let { file, body } = model;

  // Read the Excel file from the buffer
  let workbook = xlsx.read(
    file,
    { type: "buffer" },
    { cellDates: true },
    { cellNF: false },
    { cellText: false }
  );

  let main = [];
  const worksheet = workbook.Sheets[body.sheetName];

  for (let i = 1; i <= 1000; i++) {
    if (worksheet[`B${i}`] !== undefined) {
      if (worksheet[`B${i}`].v >= 1) {
        if (
          (worksheet[`E${i}`].v === 0 || worksheet[`E${i}`] === undefined) &&
          (worksheet[`G${i}`].v === 0 || worksheet[`G${i}`] === undefined)
        ) {
          break;
        }

        let info = {
          SL: worksheet[`B${i}`].v,
          issueDate: new Date(1899, 12, worksheet[`C${i}`].v + 1),
          productName: worksheet[`D${i}`].v,
          productAmount: worksheet[`E${i}`].v,
          salePrice: worksheet[`F${i}`].v,
          organizationName: worksheet[`G${i}`].v,
          address: worksheet[`H${i}`].v,
        };
        main.push(info);
      }
    }
  }

  return main;
};

module.exports = excellParser;
