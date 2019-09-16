import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
const EXCEL_TYPE =
'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8';
const EXCEL_EXT = '.xlsx';

@Injectable()
export class ExporterService {

  constructor() { }

  exportToExcel(json: any[], excelFileName: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    worksheet['!margins'] = {left: 1.0, right: 1.0, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5};

    const range: any = {
      s: {
        r: 0,
        c: 0
      },
      e: {
        r: json.length,
        c: 15
      }
    };

    for (let C = range.s.c; C < range.e.c; ++C) {
      const cell = {c: C , r: 0};
      // tslint:disable-next-line: variable-name
      const cell_ref = XLSX.utils.encode_cell(cell);
      // tslint:disable-next-line: variable-name
      const cell_address = { t: 's', v: worksheet[cell_ref].v, s: {
        font: {sz: 14, bold: true, color: '#000' }
      }};
      worksheet[cell_ref] = cell_address;
    }

    const rowInfo: XLSX.RowInfo[] = [];

    for (let R = range.s.r; R <= range.e.r; ++R) {
      rowInfo.push({ hpt: 20 });
    }

    const colInfo: XLSX.ColInfo[] = [
      { width: 8 }, { width: 12 }, { width: 45 },
      { width: 30 }, { width: 16 }, { width: 45 },
      { width: 10 }, { width: 15 }, { width: 15 },
      { width: 15 }, { width: 8 }, { width: 10 },
      { width: 10 }, { width: 10 }, { width: 12 }
    ];

    worksheet['!cols'] = colInfo;
    worksheet['!rows'] = rowInfo;

    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ['data']
    };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    // call method buffer and filename
    this.saveAsExcel(excelBuffer, excelFileName);
  }

  readFileExcel(event?: any): any {
    console.log('event ', event);
    const files = event.target.files;
    const file: File = files[0];
    const reader: FileReader = new FileReader();

    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      const workbook: XLSX.WorkBook = XLSX.read(data, { type: 'array' });
      console.log('workbook ', workbook);
      /* DO SOMETHING WITH workbook HERE */
      const worksheet: XLSX.WorkSheet = workbook.Sheets;
      console.log('worksheet ', worksheet);

      const json = XLSX.utils.sheet_to_json(worksheet.data);
      console.log('json ', json);
    };
    reader.readAsArrayBuffer(file);
  }

  private saveAsExcel(buffer: any, filename: string): void {
    const data: Blob = new Blob([buffer], {type: EXCEL_TYPE});
    FileSaver.saveAs(data, filename + '_export_' + new Date().getTime()  + EXCEL_EXT);

  }
}
