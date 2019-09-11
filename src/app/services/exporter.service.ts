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
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet} ,
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
