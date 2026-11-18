import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import type { GroupedPhysicalInventoryRow, PhysicalInventoryHeaderRow } from "../../physical-inventory-management/types";

type ExportAuditSheetArgs = {
    header: PhysicalInventoryHeaderRow;
    groupedRows: GroupedPhysicalInventoryRow[];
    branchName: string;
    supplierName: string;
    priceTypeName: string;
};

export async function exportAuditSheetToExcel({ header, groupedRows, branchName, supplierName, priceTypeName }: ExportAuditSheetArgs) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "VOS Web";
    workbook.subject = `Physical Inventory Audit Sheet ${header.ph_no ?? ""}`;
    const sheet = workbook.addWorksheet("Audit Sheet", { views: [{ state: "frozen", ySplit: 7 }] });

    sheet.columns = [
        { header: "DESCRIPTION", key: "description", width: 48 },
        { header: "SYSTEM", key: "system", width: 28 },
        { header: "PHYSICAL COUNT", key: "physicalCount", width: 28 },
        { header: "BOX", key: "box", width: 14 },
        { header: "PACK", key: "pack", width: 14 },
        { header: "TIE", key: "tie", width: 14 },
        { header: "PIECES", key: "pieces", width: 14 },
    ];

    sheet.mergeCells("A1:G1");
    sheet.getCell("A1").value = "PHYSICAL INVENTORY AUDIT SHEET";
    sheet.getCell("A1").font = { bold: true, size: 16 };
    sheet.getCell("A1").alignment = { horizontal: "center" };
    sheet.mergeCells("A2:G2");
    sheet.getCell("A2").value = header.ph_no ?? "";
    sheet.getCell("A2").alignment = { horizontal: "center" };
    sheet.getCell("A4").value = "Branch"; sheet.getCell("B4").value = branchName;
    sheet.getCell("D4").value = "Stock Type"; sheet.getCell("E4").value = header.stock_type ?? "";
    sheet.getCell("A5").value = "Supplier"; sheet.getCell("B5").value = supplierName;
    sheet.getCell("D5").value = "Price Type"; sheet.getCell("E5").value = priceTypeName;
    sheet.getRow(7).values = ["DESCRIPTION", "SYSTEM", "PHYSICAL COUNT", "BOX", "PACK", "TIE", "PIECES"];
    sheet.getRow(7).font = { bold: true, color: { argb: "FFFFFFFF" } };
    sheet.getRow(7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF333333" } };
    sheet.getRow(7).alignment = { horizontal: "center" };

    const fmtQty = (value: number) => value.toLocaleString("en-PH", { maximumFractionDigits: 0 });
    const findUnit = (row: GroupedPhysicalInventoryRow, tokens: string[], unitCountOne = false) =>
        row.rows.find((child) => {
            const name = `${child.unit_name ?? ""} ${child.unit_shortcut ?? ""}`.toUpperCase();
            return (unitCountOne && child.unit_count === 1) || tokens.some((token) => name.includes(token));
        });

    const brands = new Map<string, GroupedPhysicalInventoryRow[]>();
    for (const row of groupedRows) {
        const name = row.brand_name?.trim() || "UNBRANDED";
        brands.set(name, [...(brands.get(name) ?? []), row]);
    }
    for (const brand of [...brands.keys()].sort((a, b) => a.localeCompare(b))) {
        const heading = sheet.addRow([brand.toUpperCase()]);
        sheet.mergeCells(`A${heading.number}:G${heading.number}`);
        heading.font = { bold: true, color: { argb: "FFFFFFFF" } };
        heading.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF222222" } };
        for (const group of brands.get(brand) ?? []) {
            const system = group.rows.filter((child) => child.system_count !== 0)
                .map((child) => `${fmtQty(child.system_count)} ${child.unit_shortcut || child.unit_name || "PCS"}`)
                .join("\n") || "0";
            const physicalCount = group.rows
                .map((child) => `${fmtQty(child.physical_count)} ${child.unit_shortcut || child.unit_name || "PCS"}`)
                .join("\n");
            const box = findUnit(group, ["BOX", "CASE", "BX"]);
            const pack = findUnit(group, ["PACK", "PCK", "PK", "IB", "INNER"]);
            const tie = findUnit(group, ["TIE"]);
            const piece = findUnit(group, ["PIECE", "UNIT", "PCS"], true);
            const row = sheet.addRow([group.base_product_name, system, physicalCount, box?.unit_shortcut || box?.unit_name || "", pack?.unit_shortcut || pack?.unit_name || "", tie?.unit_shortcut || tie?.unit_name || "", piece?.unit_shortcut || piece?.unit_name || ""]);
            row.height = Math.max(20, physicalCount.split("\n").length * 15, system.split("\n").length * 15);
            row.alignment = { vertical: "top", wrapText: true };
        }
    }
    sheet.eachRow((row) => row.eachCell((cell) => {
        cell.border = { top: { style: "thin", color: { argb: "FFBBBBBB" } }, bottom: { style: "thin", color: { argb: "FFBBBBBB" } }, left: { style: "thin", color: { argb: "FFBBBBBB" } }, right: { style: "thin", color: { argb: "FFBBBBBB" } } };
    }));

    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `Physical-Inventory-Audit-Sheet-${(header.ph_no || "export").replace(/[^a-z0-9-_]/gi, "-")}.xlsx`;
    saveAs(new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
}
