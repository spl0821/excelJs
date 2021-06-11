import { Component } from "@angular/core";
import { Workbook } from "exceljs";

import * as fs from "file-saver";
import { FormGroup, FormControl } from '@angular/forms';

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  title = "angular-excel-example";
  json_data = [
    {
      name: "Raja",
      age: 20,
    },
    {
      name: "Mano",
      age: 40,
    },
    {
      name: "Tom",
      age: 40,
    },
    {
      name: "Devi",
      age: 40,
    },
    {
      name: "Mango",
      age: 40,
    },
  ];
  columnsForm = new FormGroup({
    firstColumn: new FormControl(''),
    lastColumn: new FormControl(''),
  });
  dataForm = new FormGroup({
    firstData: new FormControl(''),
    lastData: new FormControl(''),
  });
  

  /**
   *Descargar documento de excel básico
   *
   * @memberof AppComponent
   */
  async excelNormal(){
    // creación del libro de trabajo (excel)
    let workbook = new Workbook();
    // create new sheet
    let worksheet = workbook.addWorksheet("Normal Data");


    // create new sheet with color name
    // const worksheet = workbook.addWorksheet('My Sheet', {properties:{tabColor:{argb:'FFC0000'}}});

    // create new sheet with properties (color name)
    // const worksheet = workbook.addWorksheet('sheet', {properties:{tabColor:{argb:'FF00FF00'}}});

    // Create worksheets with headers and footers
    // var worksheet = workbook.addWorksheet('sheet', {
    //   headerFooter:{firstHeader: "Hello Exceljs", firstFooter: "Hello World"}
    // });

    // Set an auto filter from A1 to C1
    // worksheet.autoFilter = {
    //   from: 'A1',
    //   to: 'C1',
    // }

    // merge a range of cells
    // worksheet.mergeCells('A4:B5');

    // ... merged cells are linked
    // worksheet.getCell('B5').value = 'Hello, World!';

    // Specify Cell must be a whole number that is not 5.
    // Show the user an appropriate error message if they get it wrong
    // worksheet.getCell('A1').dataValidation = {
    //   type: 'whole',
    //   operator: 'notEqual',
    //   showErrorMessage: true,
    //   formulae: [5],
    //   errorStyle: 'error',
    //   errorTitle: 'Five',
    //   error: 'The value must not be Five'
    // };

    // Set Row 2 to Comic Sans.
    // worksheet.getRow(2).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };

    // for the wannabe graphic designers out there
    // worksheet.getCell('A1').font = {
    //   name: 'Comic Sans MS',
    //   family: 4,
    //   size: 16,
    //   underline: true,
    //   bold: true
    // };

    // for the graduate graphic designers...
    // worksheet.getCell('A2').font = {
    //   name: 'Arial Black',
    //   color: { argb: 'FF00FF00' },
    //   family: 2,
    //   size: 14,
    //   italic: true
    // };

    // Worksheets can be protected from modification by adding a password.
    // await worksheet.protect('the-password', {selectLockedCells: false});

    // fill A2 to A10 with ascending count starting from A1
    // worksheet.fillFormula('A2:A10', 'A1+1', [2,3,4,5,6,7,8,9,10]);

    let header=["Name","Age"]
    let headerRow = worksheet.addRow(header);
    for (let x1 of this.json_data)
      {
        let x2=Object.keys(x1);
        let temp=[]
        for(let y of x2)
        {
          temp.push(x1[y])
        }
        worksheet.addRow(temp)
      }
    let fname="DocumentoEnBlanco"

    //add data and file name and download
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, fname+'-'+new Date().valueOf()+'.xlsx');
    });
  }

  /**
   *Descargar documento de excel con nombres de columnas insertadas por el usuario
   *
   * @memberof AppComponent
   */
  dynamicColumns() {
    // creación del libro de trabajo (excel)
    let workbook = new Workbook();
    // create new sheet
    let worksheet = workbook.addWorksheet("Dynamic Columns");
    let header=[this.columnsForm.value.firstColumn,this.columnsForm.value.lastColumn]
    worksheet.columns = [
      {
        header: this.columnsForm.value.firstColumn, key: this.columnsForm.value.firstColumn, width: this.columnsForm.value.firstColumn.length
      },
      {
        header: this.columnsForm.value.lastColumn, key: this.columnsForm.value.lastColumn, width: this.columnsForm.value.lastColumn.length
      }
    ]
    // let headerRow = worksheet.addRow(header);
    // worksheet.properties.defaultColWidth;
    let fname="DocumentoDinamico"
    //add data and file name and download
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, fname+'-'+new Date().valueOf()+'.xlsx');
    });
  }

  /**
   *Descargar documento que contiene operaciones
   *
   * @memberof AppComponent
   */
  dataValidations(){
    // creación del libro de trabajo (excel)
    let workbook = new Workbook();
    // create new sheet
    let worksheet = workbook.addWorksheet("Data validations");
    worksheet.columns = [
      {header: 'Dato 1', key: 'dato1', width: 20},
      {header: 'Dato 2', key: 'dato2', width: 20}, 
      {header: 'Suma', key: 'suma', width: 20},
      {header: 'Resta', key: 'resta', width: 20},
      {header: 'Multiplicación', key: 'multiplicacion', width: 20},
      {header: 'División', key: 'division', width: 20},
      {header: 'Menor', key: 'menor', width: 20},
      {header: 'Mayor', key: 'mayor', width: 20},
     ];
     let number_1 = (Number)(this.dataForm.value.firstData)
     let number_2 = (Number)(this.dataForm.value.lastData)
     let suma = number_1 + number_2;
     let resta = number_1 - number_2;
     let multiplicacion = number_1 * number_2;
     let division = number_1 / number_2
     let menor;
     let mayor;
     if(number_1 > number_2){
       mayor = number_1;
       menor = number_2
     }else if(number_1 < number_2){
        mayor = number_2;
        menor = number_1
     }else{
       mayor = "iguales"
       menor = "iguales"
     }
     worksheet.addRow({dato1: this.dataForm.value.firstData, dato2: this.dataForm.value.lastData, suma: suma, resta: resta, multiplicacion: multiplicacion, division: division, menor: menor, mayor: mayor});
    let fname="DocumentoDinamico"
    //add data and file name and download
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, fname+'-'+new Date().valueOf()+'.xlsx');
    });
  }
}
