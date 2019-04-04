import { Component, OnInit } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'Excel';
  List;

  constructor(private excelService: ExcelService) {

  }

  ngOnInit() {

  }

  ExportExcel1() {
      this.excelService.generateExcel();
    }
  
  
}