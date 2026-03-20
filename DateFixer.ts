function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getUsedRange();
    let values = range.getValues();
    let rowCount = range.getRowCount();

    let lastValidDate: Date = null;

    for (let i = 0; i < rowCount; i++) {
        let dateStr = values[i][0] as string; // Column A (Date)
        let descStr = values[i][3] as string; // Column D (Description)
        let identifiedDate: Date = null;

        // 1. Transaction Code (E-code) eka check karanawa - Mekai accurate ma
        if (descStr && descStr.toString().includes("E")) {
            let eIndex = descStr.toString().indexOf("E");
            let code = descStr.toString().substring(eIndex + 1, eIndex + 7); // YYMMDD gannawa

            if (/^\d+$/.test(code)) {
                let yy = parseInt("20" + code.substring(0, 2));
                let mm = parseInt(code.substring(2, 4));
                let dd = parseInt(code.substring(4, 6));
                identifiedDate = new Date(yy, mm - 1, dd);
            }
        }

        // 2. E-code eka naththan, sequence eka check karanawa
        if (!identifiedDate && dateStr && typeof dateStr === "string") {
            if (dateStr.includes("/")) {
                let parts = dateStr.split("/");
                identifiedDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
            } else if (dateStr.includes("-")) {
                let parts = dateStr.split("-");
                let p0 = parseInt(parts[0]);
                let p1 = parseInt(parts[1]);
                let p2 = parseInt(parts[2]);

                let opt1 = new Date(p0, p1 - 1, p2); // YYYY-MM-DD
                let opt2 = new Date(p0, p2 - 1, p1); // YYYY-DD-MM (Confused format)

                if (lastValidDate) {
                    // Kalin transaction ekata chronologically lagama date eka select karanawa
                    let diff1 = Math.abs(opt1.getTime() - lastValidDate.getTime());
                    let diff2 = Math.abs(opt2.getTime() - lastValidDate.getTime());
                    identifiedDate = (diff1 < diff2) ? opt1 : opt2;
                } else {
                    identifiedDate = opt1;
                }
            }
        }

        // 3. Row eka update karanawa
        if (identifiedDate && !isNaN(identifiedDate.getTime())) {
            lastValidDate = identifiedDate;
            sheet.getCell(i, 0).setValue(identifiedDate.toLocaleDateString('en-CA')); // Force YYYY-MM-DD
            sheet.getCell(i, 1).setValue(identifiedDate.toLocaleDateString('en-CA'));
        }
    }

    // Final display format
    sheet.getRange("A:B").setNumberFormatLocal("yyyy-mm-dd");
}
