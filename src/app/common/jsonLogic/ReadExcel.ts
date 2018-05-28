import { Excelconfig as config1 } from './Excelconfig';
import * as excel from 'exceljs';

export class ReadExcel {

    cfg = new config1();

    public ReadXLS() {
        const wb = new excel.Workbook();
        const shName: string = this.cfg.getSheetName();
        let stopProcessing = true;
        let data: string[][] = [];

        wb.xlsx.readFile(this.cfg.getFilePath()).then(function() {
            const sh = wb.getWorksheet(shName);
            var fileName: string = '';

            for (let i = 2; (i <= sh.rowCount); i++) {
                let tempdata: string = this.cfg.readCellValue(i, sh, 8);
                if(tempdata.length > 0) {
                  let dataElement = tempdata.split('&&&');
                  if(i == 2) {
                    fileName = dataElement[0];
                  }
                  if(fileName == dataElement[0]) {
                    data.push(dataElement);
                  } else {
                    fileName = dataElement[0];
                    this.cfg.processFile(data);
                    data = []
                    data.push(dataElement)
                  }
                } else {
                  if(data.length > 0) {
                    this.cfg.processFile(data);
                    data = []
                  }
                }
            }
        });
    }

}




