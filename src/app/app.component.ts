import { HttpClient } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import * as fs from 'file-saver';
import { DownExcelService } from './down-excel.service';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'JsonToXl';
  // baseUrl :string ="http://localhost:8080/export/excel";
  constructor(private excel: DownExcelService){};

 ngOnInit(){};
  
DownloadExcelService()
{
  this.excel.exportExcel().subscribe(x => {
    const blob= new Blob([x],{ type:'application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
  //  if(window.navigator && window.navigator.msSaveOrOpenBlob){
  //     window.navigator.msSaveOrOpenBlob(blob);
  //    return;
  //  }
  //  fs.saveAs(blob,'EmployeeExcel.xlsx');


    const data=window.URL.createObjectURL(blob);
   const link=document.createElement('a');
   link.href=data;
   link.download='EmployeeExcel.xlsx';
   link.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,view:window}));
  
  });
}
}