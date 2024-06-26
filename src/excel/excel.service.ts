import { Injectable } from '@nestjs/common';
import * as moment from 'moment-timezone';

import { log } from 'console';
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const xlsxPopulate = require('xlsx-populate');

@Injectable()
export class ExcelService {
  locationOfRow = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
   getCurrentDateTimeInVietnam(): string {
    // Lấy thời gian hiện tại theo múi giờ Việt Nam
    const vietnamTime = moment.tz('Asia/Ho_Chi_Minh');
  
    // Định dạng ngày giờ theo yêu cầu
    const formattedTime = vietnamTime.format('H [giờ] mm [phút ngày] D/M/YYYY');
  
    return formattedTime;
  }

editCell( sheetName: string, colum: number,attendances: any[]) {
  try{
    const filePath ="C:\\Users\\Admin\\OneDrive\\Máy tính\\attendances.xlsx";
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    attendances.forEach((attendance) => {
      const cellAddress = `${this.locationOfRow[colum]}${attendance.row}`;
      log('cellAddress', cellAddress)
      if(attendance.attendance){
        sheet[cellAddress] = {v: 'Present', t: 's'};
      }else{
        sheet[cellAddress] = {v: 'Absent', t: 's'};
      }
    });
    const cellStatus = `${this.locationOfRow[colum]}${attendances.length+2}`
    const datetime = this.getCurrentDateTimeInVietnam();
    sheet[cellStatus] = {v: `Đã điểm danh${datetime}`, t: 's'};
    xlsx.writeFile(workbook, filePath);
    console.log('Cell updated successfully!');
    return true
  }catch(err){
    console.log('Error updating cell: ', err);
    return false
  }
}
readCell(sheetName, column, row) {
  try {
    const filePath = "C:\\Users\\Admin\\OneDrive\\Máy tính\\attendances.xlsx";
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    let attendances = [];
    for (let i = 1; i <= column; i++) {
      const cellAddress = `${this.locationOfRow[i]}${row}`;
      const cell = sheet[cellAddress];
      if (cell) {
        if(cell.v === 'Đã điểm danh'){
          log('Đã điểm danh',cell.v)
          attendances.push({column: i,location: cellAddress, attendance: true})
        }else{
          attendances.push({column: i,location: cellAddress, attendance: false})
        }
      }else{
        attendances.push({column: i,location: cellAddress, attendance: false})
      
      }
    }
    return attendances;
  } catch (err) {
    console.log('Error reading cell: ', err);
    return [];
  }
}

}
