import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { MongooseModule } from '@nestjs/mongoose';
import { ExcelModule } from './excel/excel.module';

@Module({
  imports: [
    MongooseModule.forRoot('mongodb+srv://nguyenminhkhoa1311:13112002@cluster0.rqtcddh.mongodb.net/'),

    ExcelModule,
  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
