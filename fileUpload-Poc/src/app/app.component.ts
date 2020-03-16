import { Component } from '@angular/core';  
import * as XLSX from 'xlsx';  
import * as FileSaver from 'file-saver';  
@Component({  
  selector: 'app-root',  
  templateUrl: './app.component.html',  
  styleUrls: ['./app.component.css']  
})  
export class AppComponent {  
  title = 'read-excel-in-angular8';  
  storeData: any;  
  csvData: any;  
  jsonData: any;  
  textData: any;  
  htmlData: any;  
  fileUploaded: File;  
  worksheet: any;  
  showData:any;
  uploadedFile(event) {  
    this.fileUploaded = event.target.files[0];  
    this.readExcel();  
  }  
  readExcel() {  
    let readFile = new FileReader();  
    readFile.onload = (e) => {  
      this.storeData = readFile.result;  
      var data = new Uint8Array(this.storeData);  
      console.log("data length :: " + data.length)
      var arr = new Array();  
      for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);  
      var bstr = arr.join("");  
      var workbook = XLSX.read(bstr, { type: "binary" });  
      workbook.SheetNames.forEach(sheetName =>{
        console.log("SheetName : " + sheetName)
      this.worksheet = workbook.Sheets[sheetName];  
      console.log(XLSX.utils.sheet_to_json(this.worksheet,{raw:true}));
      this.showData = XLSX.utils.sheet_to_json(this.worksheet,{raw:true});
      console.log("first row : " , XLSX.utils.decode_row("1")); 
  //     this.showData.forEach(element => {
  //      //let a:any[] = element.split(",");
  // console.log(element)
  //     });
      });  
    }
    
   
    readFile.readAsArrayBuffer(this.fileUploaded);  
  }  









  readAsCSV() {  
    this.csvData = XLSX.utils.sheet_to_csv(this.worksheet);  
    const data: Blob = new Blob([this.csvData], { type: 'text/csv;charset=utf-8;' });  
    FileSaver.saveAs(data, "CSVFile" + new Date().getTime() + '.csv');  
  }  
 
}  