import { excel2Svg } from "./src/excelToSvg";
import { readFileSync, writeFileSync } from "fs";

const file = readFileSync("./Layout actual play city.xlsx.xlsx");

excel2Svg(file, 'Propuesta 09-03-22').then((svg) => {
    writeFileSync("./output.svg", svg);
}).catch((err) => {
    console.error(err);
});