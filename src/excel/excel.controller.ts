import { Controller, Get, Post, Body, Patch, Param, Delete, Query, Res, HttpStatus } from '@nestjs/common';
import { Response } from 'express';
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
@Get('download')
  downloadExcel(@Res() res: Response) {
    try {
      const file = this.excelService.getExcelFile();
      res.set({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename=attendances.xlsx',
        'Content-Length': file.length,
      });
      res.status(HttpStatus.OK).send(file);
    } catch (err) {
      res.status(HttpStatus.INTERNAL_SERVER_ERROR).send('Error downloading the file');
    }
  }
}


