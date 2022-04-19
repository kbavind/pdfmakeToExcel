import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
// import { ApiService } from './api.service';
import { ExcelConverter } from 'pdfmake-to-excel';
import * as FileSaver from 'file-saver';

import * as pdfMake from "pdfmake/build/pdfmake";
import * as pdfFonts from 'pdfmake/build/vfs_fonts';
(<any>pdfMake).vfs = pdfFonts.pdfMake.vfs;
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'TestPdfExcel';
  // docDefinition: any;
  contentDefinition: any;

  constructor(private router: Router) { }
  ngOnInit() {
  }

  // content: any = [
  //   {
  //     no: 1,
  //     date: "04/01/2022",
  //     invNo: "DP1INV-22-0003",
  //     yourRef: " ",
  //     docNo: "DP1GPM4SA2103-0003/1",
  //     customer: "Excellent Sdn Bhd",
  //     acctCode: "DP1ARBILL21-0027",
  //     creditTerm: "30 Days Credit Term",
  //     due: " ",
  //     billCycle: "Monthly",
  //     handler: "M4Support01 LCS & Adila",
  //     nett: "195.60",
  //     paid: "0.00",
  //     balance: "195.60"
  //   },
  //   {
  //     no: 2,
  //     date: "06/01/2022",
  //     invNo: "DP1NV-22-0004",
  //     yourRef: "978767",
  //     docNo: " ",
  //     customer: "YYYY Sdn Bhd",
  //     acctCode: "DP-1912-0002",
  //     creditTerm: "30 Days Credit Term",
  //     due: "06/01/2022",
  //     billCycle: " ",
  //     handler: "Admin01 ZZZZ",
  //     nett: "14.91",
  //     paid: "0.00",
  //     balance: "14.91"
  //   },
  //   {
  //     no: 3,
  //     date: "06/01/2022",
  //     invNo: "DP1NV-22-0005",
  //     yourRef: "78565",
  //     docNo: " ",
  //     customer: "YYYY Sdn Bhd",
  //     acctCode: "DP-1912-0002",
  //     creditTerm: "30 Days Credit Term",
  //     due: "06/01/2022",
  //     billCycle: " ",
  //     handler: "Admin01 ZZZZ",
  //     nett: "14.91",
  //     paid: "0.00",
  //     balance: "14.91"
  //   }
  // ]


  // buildTableBody(data: any[], columns: any[]) {
  //   var body = [];

  //   body.push(columns);

  //   data.forEach(function (row) {
  //     var dataRow: any[] = [];

  //     columns.forEach(function (column) {
  //       dataRow.push(row[column].toString());
  //     });

  //     body.push(dataRow);
  //   });

  //   return body;
  // }

  // table(data: any[], columns: any[]) {
  //   return {
  //     style: 'tablecontent',
  //     table: {
  //       // style: 'tablecontent',
  //       widths: [15, 25, 40, 40, 70, 55, 60, 50, 25, 45, 50, 30, 30, 35],
  //       headerRows: 1,
  //       body: this.buildTableBody(data, columns),
  //     },
  //   };
  // }


  // pdf generation from json data and table generation

  // pdf() {
  //   this.docDefinition = {
  //     content: [
  //       { text: 'PDF Generate', style: 'header' },
  //       this.table(this.content , [  "no", "date", "invNo", "yourRef", "docNo",
  //       "customer", "acctCode", "creditTerm", "due",
  //       "billCycle", "handler","nett", "paid", "balance"]),
  //     ],
  //     styles: {
  //       header: {
  //         fontSize: 15,
  //         bold: true
  //       },
  //       tablecontent:{
  //         fontSize: 8,
  //       }
  //     }
  //   };
  //   pdfMake.createPdf(this.docDefinition).open();
  // }


  // export to excel using file saver
  
  // exportExcel() {
  //   // console.log("Exporting to Excel...");
  //   // const contentDefinition = this.getcontentDefinition();
  //   if (this.content.length > 0) {
  //     import("xlsx").then(xlsx => {
  //       const worksheet = xlsx.utils.json_to_sheet(this.content);
  //       const workbook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
  //       const excelBuffer: any = xlsx.write(workbook, { bookType: 'xlsx', type: 'array' });
  //       this.saveAsExcelFile(excelBuffer, "ExportExcel");
  //     });
  //   }
  // }
  // saveAsExcelFile(buffer: any, fileName: string): void {
  //   let EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
  //   let EXCEL_EXTENSION = '.xlsx';
  //   const data: Blob = new Blob([buffer], {
  //     type: EXCEL_TYPE
  //   });
  //   FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
  // }


  // pdf generation from pdfmake content
  getcontentDefinition() {
    this.contentDefinition = {
      pageOrientation: 'landscape',
      content: [
        { text: 'Sales Invoice Listing', style: 'title' },
        { text: 'Period: 1 Jan 2022 to 31 Dec 2022\n\n', style: 'subtitle' },
        { text: 'MYR', style: 'title' },
        {
          style: 'tableExample',
          table: {
            widths: [15, 25, 40, 40, 70, 55, 60, 50, 25, 45, 50, 30, 30, 35],
            headerRows: 1,
            body: [
              [{ text: 'No.', style: 'tableheader' }, { text: 'Date', style: 'tableheader' }, { text: 'Inv Date', style: 'tableheader' }, { text: 'Your Ref', style: 'tableheader' }, { text: 'Doc No.', style: 'tableheader' }, { text: 'Customer', style: 'tableheader' }, { text: 'Acct Code', style: 'tableheader' }, { text: 'Credit Term', style: 'tableheader' }, { text: 'Due', style: 'tableheader' }, { text: 'BillCycle', style: 'tableheader' }, { text: 'Handler', style: 'tableheader' }, { text: 'Nett', style: 'tableheader' }, { text: 'Paid', style: 'tableheader' }, { text: 'Balance', style: 'tableheader' }],
              [{ text: "1." }, { text: "04/01/2022" }, { text: "DP1NV-22-0003" }, { text: " " }, { text: "DP1GP-M4SA2103-0003/1" }, { text: "Excellent Sdn Bhd" }, { text: "DP1ARBILL21-0027" }, { text: "30 Days Credit Term" }, { text: " " }, { text: "Monthly" }, { text: "M4Support01 LCS,Adil" }, { text: "195.60" }, { text: "0.00" }, { text: "195.60" }],
              [{ text: "2." }, { text: "06/01/2022" }, { text: "DP1NV-22-0004" }, { text: "978767" }, { text: " " }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: "06/01/2022" }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "3." }, { text: "06/01/2022" }, { text: "DP1NV-22-0005" }, { text: "78565" }, { text: " " }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: "06/01/2022" }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "4." }, { text: "08/02/2022" }, { text: "DP1NV-22-0006" }, { text: " " }, { text: "DP1GP-CSDA2107-0005/1" }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: " " }, { text: "Monthly" }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "5." }, { text: "12/02/2022" }, { text: "DP1NV-22-0007" }, { text: "YourRef(ORDR)" }, { text: "DP1GP-LCS02101-0003-01/1" }, { text: "XXXX XX SDN BHD" }, { text: "DP1-2101-0008" }, { text: "30 Days Credit Term" }, { text: " " }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "562.89" }, "0.00", "562.89"],
              [{ text: "6." }, { text: "18/02/2022" }, { text: "DP1NV-22-0009" }, { text: " " }, { text: " " }, { text: "XXXX XX SDN BHD" }, { text: "DP1-2101-0008" }, { text: "30 Days Credit Term" }, { text: " " }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "66.00" }, { text: "0.00" }, { text: "66.00" }],
              [{ text: "7." }, { text: "21/02/2022" }, { text: "DP1NV-22-0010" }, { text: " " }, { text: "DP1GP-P3FA2201-0005/1" }, { text: "XXXX XX SDN BHD" }, { text: "DP1-2101-0008" }, { text: "30 Days Credit Term" }, { text: " " }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "1416.96" }, { text: "0.00" }, { text: "1416.96" }],
              [{ text: "8." }, { text: "18/02/2022" }, { text: "DP1NV-22-0011" }, { text: " " }, { text: "DP1GP-DILA2112-0024/1" }, { text: "XXXX XX SDN BHD" }, { text: "DP1-2101-0008" }, { text: "3 Days Credit Term" }, { text: "05/03/2022" }, { text: " " }, { text: "abcdef" }, { text: "151.07" }, { text: "0.00" }, { text: "151.07" }],
              [{ text: "9." }, { text: "06/03/2022" }, { text: "DP1NV-22-0012" }, { text: " " }, { text: "DP1GP-CSDA2107-0003/1" }, { text: "test" }, { text: "C-10001" }, { text: "30 Days Credit Term" }, { text: "05/04/2022" }, { text: "Monthly" }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "10." }, { text: "08/03/2022" }, { text: "DP1NV-22-0013" }, { text: " " }, { text: "DP1GP-CSDA2107-0005/1" }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: "07/04/2022" }, { text: "Monthly" }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "11." }, { text: "01/04/2022" }, { text: "DP1INV-22-0008" }, { text: " " }, { text: "DP1GP-M4SA2202-0003/1" }, { text: "Excellent Sdn Bhd" }, { text: "DP1ARBILL21-0027" }, { text: "30 Days Credit Term" }, { text: " " }, { text: "Monthly" }, { text: "qwertyuiop" }, { text: "10.22" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "12." }, { text: "14/02/2022" }, { text: "INV-22-04-0001" }, { text: " " }, { text: "DP1DO-22-0001" }, { text: "Excellent Sdn Bhd" }, { text: "DP1ARBILL21-0027" }, { text: "30 Days Credit Term" }, { text: " " }, { text: " " }, { text: "qwertyuiop" }, { text: "100.00" }, { text: "0.00" }, { text: "100.00" }],
              [{ text: "13." }, { text: "01/04/2022" }, { text: "INV-22-04-0003" }, { text: " " }, { text: "DP1DS-V522011-0001/1" }, { text: "Asraf" }, { text: "DP-1911-0089" }, { text: "Cash on Delivery" }, { text: " " }, { text: "Monthly" }, { text: "abcdef" }, { text: "5000.00" }, { text: "0.00" }, { text: "5000.00" }],
              [{ text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: 'TOTAL', bold: false, style: 'tablefooter' }, { text: '7577.29', style: 'tablefooter' }, { text: '0.00', style: 'tablefooter' }, { text: '7,567.07', style: 'tablefooter' }]
            ]
          },
          layout: 'lightHorizontalLines',
        },
        { text: '\n' },
        {
          style: 'tableExample',
          table: {
            widths: [15, 25, 40, 40, 70, 55, 60, 50, 25, 45, 50, 30, 30, 35],
            headerRows: 1,
            body: [
              [{ text: 'No.', style: 'tableheader' }, { text: 'Date', style: 'tableheader' }, { text: 'Inv Date', style: 'tableheader' }, { text: 'Your Ref', style: 'tableheader' }, { text: 'Doc No.', style: 'tableheader' }, { text: 'Customer', style: 'tableheader' }, { text: 'Acct Code', style: 'tableheader' }, { text: 'Credit Term', style: 'tableheader' }, { text: 'Due', style: 'tableheader' }, { text: 'BillCycle', style: 'tableheader' }, { text: 'Handler', style: 'tableheader' }, { text: 'Nett', style: 'tableheader' }, { text: 'Paid', style: 'tableheader' }, { text: 'Balance', style: 'tableheader' }],
              [{ text: "1." }, { text: "04/01/2022" }, { text: "DP1NV-22-0003" }, { text: " " }, { text: "DP1GP-M4SA2103-0003/1" }, { text: "Excellent Sdn Bhd" }, { text: "DP1ARBILL21-0027" }, { text: "30 Days Credit Term" }, { text: " " }, { text: "Monthly" }, { text: "M4Support01 LCS,Adil" }, { text: "195.60" }, { text: "0.00" }, { text: "195.60" }],
              [{ text: "2." }, { text: "06/01/2022" }, { text: "DP1NV-22-0004" }, { text: "978767" }, { text: " " }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: "06/01/2022" }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "3." }, { text: "06/01/2022" }, { text: "DP1NV-22-0005" }, { text: "78565" }, { text: " " }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: "06/01/2022" }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "4." }, { text: "08/02/2022" }, { text: "DP1NV-22-0006" }, { text: " " }, { text: "DP1GP-CSDA2107-0005/1" }, { text: "YYYY Sdn Bhd" }, { text: "DP-1912-0002" }, { text: "30 Days Credit Term" }, { text: " " }, { text: "Monthly" }, { text: "Admin01 ZZZZ" }, { text: "14.91" }, { text: "0.00" }, { text: "14.91" }],
              [{ text: "5." }, { text: "12/02/2022" }, { text: "DP1NV-22-0007" }, { text: "YourRef(ORDR)" }, { text: "DP1GP-LCS02101-0003-01/1" }, { text: "XXXX XX SDN BHD" }, { text: "DP1-2101-0008" }, { text: "30 Days Credit Term" }, { text: " " }, { text: " " }, { text: "Admin01 ZZZZ" }, { text: "562.89" }, "0.00", "562.89"],
              [{ text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: ' ', style: 'tablefooter' }, { text: 'TOTAL', bold: false, style: 'tablefooter' }, { text: '7577.29', style: 'tablefooter' }, { text: '0.00', style: 'tablefooter' }, { text: '7,567.07', style: 'tablefooter' }]
            ]
          },
          layout: 'lightHorizontalLines',
        },
      ],
      styles: {
        title: {
          bold: true,
        },
        subtitle: {
          color: 'gray',
          bold: true,
          fontSize: 9
        },
        tableheader: {
          color: 'gray',
          bold: true
        },
        tableExample: {
          fontSize: 8
        },
        tablefooter: {
          fillColor: '#eeeeee',
          bold: true
        }
      }
    };
    return this.contentDefinition;
  }

  getTable = () => {
    const data = this.getcontentDefinition();

    return data.content.filter((obj: Object) =>  obj.hasOwnProperty('table'));
  };

//   getContentTable = () => {
//     this.contentDefinition.map((c: { table: { length: any; text: any; body: any; }; }, index: any) => {
//         if (c?.table?.length) {
//             return {
//                 sheetName: `sheetName-${c.table.text}`,
//                 sheetData: c.table.body || []
//             }   
//         }
//     }).filfer((l: any) => l)
// }


  // Download Pdf(data from contentDefinition pdfmake)
  generatePdf() {
    console.log("Exporting to PDF...");
    this.contentDefinition = this.getcontentDefinition();
    pdfMake.createPdf(this.contentDefinition).open();
  }

  // Download Excel(data from contentDefinition pdfmake)
  generateExcel() {
    console.log("Exporting to Excel...");
    const tables: Array<Object> = this.getTable();
    let contentD = {
      "title": "Sales Invoice Listing",
      data: []
    };

    console.log('Tables:' ,tables)

    tables.map((tableData, index) => { 
      contentD.data.push(
        // @ts-ignore
        { sheetName: 'Sheet_name ' + (index+1), sheetData: tableData.table.body} 
      )
    })
    
    const exporter = new ExcelConverter('InvoiceExport', contentD); 
    exporter.downloadExcel();
  };


}
