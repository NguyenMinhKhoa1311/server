import { Controller, Get, Post, Body, Patch, Param, Delete, Query } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { log } from 'console';


@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('edit')
  edit( @Query('sheetName') sheetName: string, @Query('colum') colum: number, @Body('attendances') attendances: any[]){
    log('edit', sheetName, colum, attendances)
    return this.excelService.editCell(sheetName,colum, attendances)
  }
  @Post('read')
  read( @Query('sheetName') sheetName: string, @Query('colum') colum: number, @Query('row') row: number){
    return this.excelService.readCell(sheetName, colum, row)
}
}


