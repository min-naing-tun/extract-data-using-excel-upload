import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'extract-data-using-excel-upload';
  public binaryString: string = "";
  public workBook: string = "";
  public workSheet: string = "";
  public excelDataValue: string = "";

  fileUploaded(event: any) {
    //Get event target for single file or multi files
    const target: DataTransfer = <DataTransfer>event.target;
    if (target.files.length > 1) {
      console.log('Multiple Files Upload Not Supported');
    } else {
      const reader: FileReader = new FileReader();
      //file reader read this file
      reader.readAsBinaryString(target.files[0]);
      reader.onload = (e: any) => {
        
        //event
        console.log(e);

        //html element by event target
        console.log(target);

        //uploaded file result data (raw)
        const bstr: string = e.target.result;
        this.binaryString = bstr;
        console.log(bstr);

        //read data the whole excel file
        //cellDates: true will be date data correctness
        const wb: XLSX.WorkBook = XLSX.read(bstr, {
          type: 'binary',
          cellDates: true,
        });
        this.workBook = JSON.stringify(wb);
        console.log(wb);

        //select and read first sheet of the whole data
        const wsname = wb.SheetNames[0];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];
        this.workSheet = JSON.stringify(ws);
        console.log(ws);

        //read selected sheet's data line by line and store data to (excelData)
        let excelData = XLSX.utils.sheet_to_json(ws, { header: 1 });
        this.excelDataValue = JSON.stringify(excelData);
        console.log(excelData);
      };
      
    }
  }
}
