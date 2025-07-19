import { CellErrorValue, CellFormulaValue, CellHyperlinkValue, CellRichTextValue, CellSharedFormulaValue, Color, FillGradientAngle, FillGradientPath, FillPattern, Workbook, Worksheet } from "exceljs";
import { js2xml } from "xml-js";
import sizeOf from "image-size";
import sanitizeHtml from 'sanitize-html';
import { getThemeColors } from "./theme parser";

const borderStyles = {
    "medium": {
        "stroke-width": 2
    },
    "thin": {},
    "dotted": {
        "stroke-dasharray": 3
    },
    "thick": {
        "stroke-width": 4
    },
    "dashDot": {
        "stroke-dasharray": '10,3,3,3'
    },
    "hair": {
        "stroke-dasharray": 2
    },
    "dashDotDot": {
        "stroke-dasharray": '10,3,3,3,3,3'
    },
    "dashed": {
        "stroke-dasharray": '4 2'
    },
    "mediumDashDotDot": {
        "stroke-dasharray": '10,3,3,3,3,3',
        "stroke-width": 2
    },
    "slantDashDot": {
        "stroke-dasharray": '10,3,3,3',
        "stroke-width": 2
    },
    "mediumDashDot": {
        "stroke-dasharray": '10,3,3,3',
        "stroke-width": 2
    },
    "mediumDashed": {
        "stroke-dasharray": '4 2',
        "stroke-width": 2
    },
    "double": {
        "stroke-width": 4
    }
}

const ratioX = 15. / 190500.
const ratioY = 10.71 * 7.45 / 762000.

// Patterns
// `darkGray        <pattern id="Pattern" x="0" y="0" width="4" height="4" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="4" height="4" fill="black"/>
// <circle cx="0" cy="0" r="0.5" fill="white"/>
// <circle cx="0" cy="4" r="0.5" fill="white"/>
// <circle cx="2" cy="2" r="0.5" fill="white"/>
// <circle cx="4" cy="0" r="0.5" fill="white"/>
// <circle cx="4" cy="4" r="0.5" fill="white"/>
// </pattern>
// mediumGray  <pattern id="Pattern" x="0" y="0" width="4" height="4" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="4" height="4" fill="black"/>
// <circle cx="0" cy="0" r="0.5" fill="white"/>
// <circle cx="0" cy="4" r="0.5" fill="white"/>
// <circle cx="2" cy="2" r="0.5" fill="white"/>
// <circle cx="4" cy="0" r="0.5" fill="white"/>
// <circle cx="4" cy="4" r="0.5" fill="white"/>
// </pattern>
// lightGray   <pattern id="Pattern" x="0" y="0" width="6" height="6" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="6" height="6" fill="black"/>
//   <circle cx="0" cy="0" r="1" fill="white"/>
//   <circle cx="0" cy="6" r="1" fill="white"/>
//   <circle cx="3" cy="3" r="1" fill="white"/>
//   <circle cx="6" cy="0" r="1" fill="white"/>
//   <circle cx="6" cy="6" r="1" fill="white"/>
// </pattern>
// gray125 <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="8" height="8" fill="black"/>
//   <circle cx="0" cy="0" r="1" fill="white"/>
//   <circle cx="0" cy="8" r="1" fill="white"/>
//   <circle cx="4" cy="4" r="1" fill="white"/>
//   <circle cx="8" cy="0" r="1" fill="white"/>
//   <circle cx="8" cy="8" r="1" fill="white"/>
// </pattern>
// gray0625    
// <pattern id="Pattern" x="0" y="0" width="10" height="10" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="10" height="10" fill="black"/>
//   <circle cx="0" cy="0" r="1" fill="white"/>
//   <circle cx="0" cy="10" r="1" fill="white"/>
//   <circle cx="5" cy="5" r="1" fill="white"/>
//   <circle cx="10" cy="0" r="1" fill="white"/>
//   <circle cx="10" cy="10" r="1" fill="white"/>
// </pattern>
// </defs>
// darkHorizontal  <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="8" height="4" fill="black"/>
//   <rect x="0" y="4" width="8" height="4" fill="white"/>
// </pattern>
// darkVertical    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="4" height="8" fill="black"/>
// <rect x="4" y="0" width="4" height="8" fill="white"/>
// </pattern>
// darkDown    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="8" height="8" fill="black"/>
//   <path d="M0 0L8 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
//   <path d="M8 0L16 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
//   <path d="M0 8L8 16" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// </pattern>
// darkUp  <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="black"/>
// <path d="M8 0L0 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// <path d="M8 8L0 16" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// <path d="M0 0L-8 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// </pattern>
// darkGrid    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="black"/>
// <path d="M4 0L4 8" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// <path d="M0 4L8 4" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// </pattern>
// darkTrellis <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="black"/>
// <path d="M8 0L0 8" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// <path d="M0 0L8 8" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// </pattern>
// lightHorizontal <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="7" fill="white"/>
// <rect x="0" y="7" width="8" height="1" fill="black"/>
// </pattern>
// lightVertical   <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="7" height="8" fill="white"/>
// <rect x="7" y="0" width="1" height="8" fill="black"/>
// </pattern>
// lightDown   <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M0 0L8 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M8 0L16 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 8L8 16" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// lightUp <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M8 0L0 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M8 8L0 16" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 0L-8 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// lightGrid   <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M4 0L4 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 4L8 4" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// lightTrellis    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M8 0L0 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 0L8 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// `
export const excel2Svg = async (file: Buffer, worksheetName?: string) => {
    const workbook = new Workbook();
    const themes = await getThemeColors(file)
    const getColor = (color?: Partial<Color & { tint: number }>): string | null => {
    if (color?.argb) {
        return "#" + color.argb.slice(2, 8);
    } else if (color?.theme != null) {
        let theme = themes[color.theme];
        const tint = color.tint;
        if (tint) {
            theme = CalculateFinalLumValue(tint, theme);
        }
        return "#" + theme;
    } else {
        return null;
    }
}
    await workbook.xlsx.load(file)
    const sheet: Worksheet =
        workbook.worksheets.find(sheet => sheet.name === worksheetName) ??
        workbook.worksheets.find(sheet => sheet.id === workbook.views?.[0]?.activeTab) ??
        workbook.worksheets[0]

    /* sheetIds for this file: 51, 71, 60, 50, 66, 67 */
    const heights: number[] = []
    for (let i = 1; i <= sheet.rowCount; i++) {
        const row = sheet.getRow(i)
        heights.push(row.height || 15);
    }
    let widths: number[] = []
    for (const column of sheet.columns) {
        widths.push((column.width || 10.71) * 7.45);
    }
    if (widths.length === 0) widths = [10.71]
    const image: {
        svg: {
            _attributes: {
                version: number;
                width: number;
                height: number;
                xmlns: string;
            };
            defs: {
                linearGradient: any[];
                radialGradient: any[];
                image: any[];
            };
            rect: any[];
            path: any[];
            text: any[];
            use: any[];
        }
    } = {
        svg: {
            _attributes: {
                version: 1.1,
                width: 20000,
                height: 20000,
                xmlns: "http://www.w3.org/2000/svg",
            },
            defs: {
                linearGradient: [],
                radialGradient: [],
                image: []
            },
            rect: [],
            path: [],
            text: [],
            use: []
        },
    };
    let top: number;
    let bottom: number;
    let left: number;
    let right: number;
    let color: string | null = "";
    let x: number;
    let y: number;
    let height: number;
    let width: number;
    let totWidth: number = 0
    let totHeight: number = 0
    const defaultPadding = 3;

    for (let i = 1; i <= sheet.rowCount; i++) {
        const row = sheet.getRow(i);
        for (let j = 1; j <= row.cellCount; j++) {
            const cell = row.getCell(j)
            if (cell.value !== undefined && cell.master !== cell) {
                //   appendFileSync('nuevo.txt', cell.address + '\n')
                continue
            };
            if (cell.isMerged) {
                const merge = sheet.model.merges.find((merge) => {
                    const splitted = merge.split(":");
                    return splitted[0] === cell.fullAddress.address;
                });

                top = row.number - 1;
                bottom = !merge ? top + 1 : sheet.getCell(merge.split(":")[1]).fullAddress.row;
                left = cell.fullAddress.col - 1;
                right = !merge ? top + 1 : sheet.getCell(merge.split(":")[1]).fullAddress.col;

            } else {
                top = row.number - 1;
                bottom = top + 1;
                left = cell.fullAddress.col - 1;
                right = left + 1;
            }
            x =
                widths.slice(0, left).length == 0
                    ? 0
                    : widths.slice(0, left).reduce((a, b) => a + b);
            y =
                heights.slice(0, top).length == 0
                    ? 0
                    : heights.slice(0, top).reduce((a, b) => { return a + b }, 0);
            height = heights.slice(top, bottom).reduce((a, b) => a + b, 0);
            width = widths.slice(left, right).reduce((a, b) => a + b, 0);
            x = Math.round(x)
            y = Math.round(y)
            height = Math.round(height)
            width = Math.round(width)
            if (x + width > totWidth) totWidth = x + width
            if (y + height > totHeight) totHeight = y + height
            // For the filling
            if (cell.fill?.type === "pattern") {
                const filling = cell.style.fill as FillPattern
                color = getColor(filling.fgColor);
            } else if (cell.style?.fill?.type === "gradient") {
                if (cell.style.fill.gradient === "angle" || cell.style.fill.gradient == null) {
                    const filling = cell.style.fill as FillGradientAngle;
                    let i = image.svg.defs.linearGradient.findIndex(
                        (x) => x.reference == cell.style.fill
                    );
                    if (i == -1) {
                        image.svg.defs.linearGradient.push({
                            _attributes: {
                                id: "linearGradient" + cell.address,
                                ...getCoords(filling.degree),
                            },
                            reference: filling,
                            stop: filling.stops.map((stop) => {
                                return {
                                    _attributes: {
                                        "stop-color": getColor(stop.color),
                                        offset: stop.position,
                                    },
                                };
                            }),
                        });
                        color = `url(#linearGradient${cell.address})`;
                    } else {
                        color = `url(#${image.svg.defs.linearGradient[i]._attributes.id})`;
                    }
                } else if (cell.style.fill.gradient == "path") {
                    const filling = cell.style.fill as FillGradientPath;
                    let i = image.svg.defs.radialGradient.findIndex(
                        (x) => x.reference == filling
                    );
                    if (i == -1) {
                        image.svg.defs.radialGradient.push({
                            _attributes: {
                                id: "radialGradient" + cell.address,
                                cx: filling.center.left,
                                cy: filling.center.top,
                                r: 1
                            },
                            reference: filling,
                            stop: filling.stops.map((stop) => {
                                return {
                                    _attributes: {
                                        "stop-color": getColor(stop.color),
                                        offset: stop.position,
                                    },
                                };
                            }),
                        });
                        color = `url(#radialGradient${cell.address})`;
                    } else {
                        color = `url(#${image.svg.defs.radialGradient[i]._attributes.id})`;
                    }
                } else {
                    color = "#FFFFFF";
                }
            } else {
                color = "#FFFFFF";
            }
            if (color) {
                image.svg.rect.push({
                    _attributes: {
                        x: x,
                        y: y,
                        height: height,
                        width: width,
                        fill: color,
                    },
                });
            }
            //For the borders
            if (cell.style.border) {
                Object.entries(cell.style.border).forEach(([side, value]) => {
                    let color = getColor(value.color);
                    color = color == null ? "black" : color;
                    let id =
                        value.style + "," + color
                    let i = image.svg.path.findIndex((x => x._attributes.id == id));
                    if (i == -1) {
                        i = image.svg.path.push({
                            _attributes: {
                                id: id,
                                d: '',
                                stroke: color,
                                fill: 'none',
                                ...value.style && { ...borderStyles[value.style] }
                            },
                        }) - 1;
                    }
                    switch (side) {
                        case "left":
                            image.svg.path[i]._attributes.d += 'M' + x + ' ' + y + 'L' + x + ' ' + (y + height)
                            break;
                        case "right":
                            image.svg.path[i]._attributes.d += 'M' + (x + width) + ' ' + y + 'L' + (x + width) + ' ' + (y + height)
                            break;
                        case "top":
                            image.svg.path[i]._attributes.d += 'M' + x + ' ' + y + 'L' + (x + width) + ' ' + y
                            break;
                        case "bottom":
                            image.svg.path[i]._attributes.d += 'M' + x + ' ' + (y + height) + 'L' + (x + width) + ' ' + (y + height)
                            break
                        case "diagonal":
                            image.svg.path[i]._attributes.d += 'M' + x + ' ' + y + 'L' + (x + width) + ' ' + (y + height)
                    }
                });
            }
            // For the text
            let value = ""
            if ((cell.value as CellErrorValue)?.error) {
                value = (cell.value as CellErrorValue).error
            } else if ((cell.value as CellFormulaValue)?.formula || (cell.value as CellSharedFormulaValue)?.sharedFormula) {
                if (((cell.value as CellFormulaValue).result as CellErrorValue)?.error) {
                    value = ((cell.value as CellFormulaValue).result as CellErrorValue).error
                } else {
                    value = cell.result?.toString()??""
                }
            } else if ((cell.value as CellRichTextValue)?.richText) {
                value = (cell.value as CellRichTextValue).richText.map(e => e.text).join('')
            } else if ((cell.value as CellHyperlinkValue)?.text) {
                value = (cell.value as CellHyperlinkValue).text
            } else {
                value = cell.value?.toString()??""
            }
            if (value != null) {
                const text: { _attributes: { x: number; y: number; "font-size"?: number | string; "text-anchor"?: string; "dominant-baseline"?: string; "font-weight"?: string; fill?: string }; _text: string; } = {
                    _attributes: {
                        x: Math.round(x + defaultPadding),
                        y: Math.round(y + height - defaultPadding),
                        "font-size": cell.style?.font?.size,
                    },
                    _text: sanitizeHtml(value),
                };
                let color = getColor(cell.style?.font?.color)
                text._attributes['fill'] = color ?? '#000000'
                if (cell.style?.alignment) {
                    switch (cell.style?.alignment.horizontal) {
                        case "right":
                            text._attributes.x = Math.round(x + width - defaultPadding);
                            text._attributes["text-anchor"] = "end";
                            break;
                        case "center":
                            text._attributes.x = Math.round(x + width / 2);
                            text._attributes["text-anchor"] = "middle";
                            break;
                    }
                    switch (cell.style?.alignment.vertical) {
                        case "top":
                            text._attributes.y = Math.round(y + defaultPadding);
                            text._attributes["dominant-baseline"] = "hanging";
                            break;
                        case "middle":
                            text._attributes.y = Math.round(y + height / 2);
                            text._attributes["dominant-baseline"] = "middle";
                            break;
                    }
                }
                if (cell.style.font?.bold)
                    text._attributes["font-weight"] = "bold";
                image.svg.text.push(text);
            }
        };
    };
    image.svg.defs.linearGradient = image.svg.defs.linearGradient.map(
        ({ reference, ...theRest }) => theRest
    );
    image.svg.defs.radialGradient = image.svg.defs.radialGradient.map(
        ({ reference, ...theRest }) => theRest
    );
    // for the images
    const images = sheet.getImages();
    for (const image1 of images) {
        const media = workbook.model.media.find((e: any) => e.index == +image1.imageId)
        if (!media || media.type != 'image') continue;

        let i = image.svg.defs.image.findIndex(
            (x) => x.id == 'image' + image1.imageId
        );
        const imagePba = Buffer.from(media.buffer)
        if (i == -1) {
            i = image.svg.defs.image.push({
                _attributes: {
                    id: 'image' + image1.imageId,
                    href: "data:image/png;base64, " + imagePba.toString('base64')
                },
            }) - 1;
        }
        const dimensions = sizeOf(imagePba)
        top = image1.range.tl.nativeRow;
        bottom = image1.range.br.nativeRow;
        left = image1.range.tl.nativeCol;
        right = image1.range.br.nativeCol;
        x =
            (widths.slice(0, left).length == 0
                ? 0
                : widths.slice(0, left).reduce((a, b) => a + b)) + image1.range.tl.nativeColOff * ratioX;
        y =
            (heights.slice(0, top).length == 0
                ? 0
                : heights.slice(0, top).reduce((a, b) => a + b)) + image1.range.tl.nativeRowOff * ratioY;
        height = (heights.slice(top, bottom).length == 0
            ? 0
            : heights.slice(top, bottom).reduce((a, b) => a + b)) + image1.range.br.nativeRowOff * ratioY - image1.range.tl.nativeRowOff * ratioY;
        width = widths.slice(left, right).length == 0
            ? 0
            : widths.slice(left, right).reduce((a, b) => a + b) + image1.range.br.nativeColOff * ratioX - image1.range.tl.nativeColOff * ratioX;
        x = Math.round(x)
        y = Math.round(y)
        height = Math.round(height)
        width = Math.round(width)
        image.svg.use.push({
            _attributes: {
                href: "#" + image.svg.defs.image[i]._attributes.id,
                transform: `matrix(${width / dimensions.width} 0 0 ${height / dimensions.height} ${x} ${y})`
            }
        })

    }
    return js2xml(image, { compact: true });

}

function CalculateFinalLumValue(tint: number, color: string): string {
    let [r, g, b] = [parseInt(color.slice(0,2), 16), parseInt(color.slice(2,4), 16), parseInt(color.slice(4), 16)];
    [r, g, b] = [r, g, b].map(el => {
        if (tint > 0) {
            el += (255 - el) * tint
        } else {
            el *= (1 + tint)
        }
        return Math.round(el)
    })
    return [r, g, b].map(e => e.toString(16)).join('')
}

function getCoords(angle: number) {
    if (angle == null) angle = 0;
    angle = (angle * 2 * Math.PI) / 360.0;
    const coords = { x1: 0.5, y1: 0.5, x2: 0, y2: 0 };
    const cuarto = Math.PI / 4;
    if (angle >= cuarto && angle < 3 * cuarto) {
        coords.y2 = 1;
        coords.x2 = (coords.y2 - coords.y1) / Math.tan(angle) + coords.x1;
    } else if (angle >= 3 * cuarto && angle < 5 * cuarto) {
        coords.x2 = 0;
        coords.y2 = (coords.x2 - coords.x1) / Math.tan(angle) + coords.y1;
    } else if (angle >= 5 * cuarto && angle < 7 * cuarto) {
        coords.y2 = 0;
        coords.x2 = (coords.y2 - coords.y1) / Math.tan(angle) + coords.x1;
    } else {
        coords.x2 = 1;
        coords.y2 = (coords.x2 - coords.x1) / Math.tan(angle) + coords.y1;
    }
    coords.y1 = 1 - coords.y2;
    coords.x1 = 1 - coords.x2;
    if (angle == 0) coords.y1 = 0, coords.y2 = 0
    return coords;
}