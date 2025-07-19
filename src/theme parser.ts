import { Open } from 'unzipper'
import { xml2js } from "xml-js";

function findElement(elements: any[], name: string) {
    return elements?.find(e => e.type === 'element' && e.name === name);
}

function findColor(clrScheme: any, name: string): string {
    const color = clrScheme?.elements?.find((e: any) => e.name === name)?.elements?.[0]?.attributes
    return color?.lastClr  ?? color?.val
}

export async function getThemeColors(xlsxPath: Buffer): Promise<string[]> {
    const zip = await Open.buffer(xlsxPath);
    const themeEntry = zip.files.find(f => f.path === "xl/theme/theme1.xml");

    if (!themeEntry) throw null;

    const xml = await themeEntry.buffer();
    const json = xml2js(xml.toString(), { compact: false });

    const aTheme = findElement(json.elements, 'a:theme');
    const themeElements = findElement(aTheme?.elements, 'a:themeElements');
    const clrScheme = findElement(themeElements?.elements, 'a:clrScheme');

    return [
        findColor(clrScheme, 'a:lt1')??"FFFFFF",
        findColor(clrScheme, 'a:dk1')??"000000",
        findColor(clrScheme, 'a:lt2')??"EEECE1",
        findColor(clrScheme, 'a:dk2')??"1F497D",
        findColor(clrScheme, 'a:accent1')??"4F81BD",
        findColor(clrScheme, 'a:accent2')??"C0504D",
        findColor(clrScheme, 'a:accent3')??"9BBB59",
        findColor(clrScheme, 'a:accent4')??"8064A2",
        findColor(clrScheme, 'a:accent5')??"4BACC6",
        findColor(clrScheme, 'a:accent6')??"F79646"
    ]
}