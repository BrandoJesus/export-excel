import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
const EXCEL_TYPE =
'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8';
const EXCEL_EXT = '.xlsx';

@Injectable()
export class ExporterService {
  rowInfo: XLSX.RowInfo[] = [];
  colInfo: XLSX.ColInfo[] = [];
  worksheet: XLSX.WorkSheet;
  workbook: XLSX.WorkBook;
  range: XLSX.Range;
  excelBuffer: Buffer;
  constructor() {}

  exportToExcel(json: any[], excelFileName: string, type?: string): void {
    // console.log('json ', json, json.length);
    this.colInfo = this.setColInfo(type);
    this.worksheet = this.formatWorkSheet(json);

    this.workbook = {
      Sheets: { data: this.worksheet },
      SheetNames: ['data']
    };
    this.excelBuffer = XLSX.write(this.workbook, { bookType: 'xlsx', type: 'array' });
    // call method buffer and filename
    this.saveAsExcel(this.excelBuffer, excelFileName);
  }

  formatWorkSheet(json: any) {
    const ws = XLSX.utils.json_to_sheet(json);
    const range: any = {
      s: { r: 0, c: 0  },
      e: { r: json.length, c: this.colInfo.length }
    };

    for (let R = range.s.r; R <= range.e.r; ++R) {
      this.rowInfo.push({ hpt: 20 });
    }

    ws['!margins'] = {left: 1.0, right: 1.0, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5};
    ws['!cols'] = this.colInfo;
    ws['!rows'] = this.rowInfo;
    return ws;
  }

  setColInfo(type: string) {
    switch (type) {
      case 'QUOTATION':
        return this.colInfo = [
          { width: 8 }, { width: 12 }, { width: 45 },
          { width: 30 }, { width: 16 }, { width: 45 },
          { width: 10 }, { width: 15 }, { width: 15 },
          { width: 15 }, { width: 8 }, { width: 10 },
          { width: 10 }, { width: 10 }, { width: 12 }
        ];
      default:
        return this.colInfo = [
          { width: 8 }, { width: 12 }, { width: 45 },
          { width: 30 }, { width: 16 }, { width: 45 },
          { width: 10 }, { width: 15 }, { width: 15 },
          { width: 15 }, { width: 8 }, { width: 10 },
          { width: 10 }, { width: 10 }, { width: 12 }
        ];
    }
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
