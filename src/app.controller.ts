import { Controller, Get, Res } from '@nestjs/common';
import { AppService } from './app.service';
// import { Readable } from 'stream';
import { Response } from 'express';
import { Screen } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  @Get('/excel')
  async getExcel(@Res() res: Response) {
    const excel = await this.appService.readExcel();
    const screen: Screen = {
      title: '页面1',
      components: [
        {
          uuid: '121231231231231',
          title: '人力分布',
          typeName: '表格',
          data: [
            [
              { name: '后台', value: 14, increase: 0, decrease: 1 },
              { name: '前端', value: 20, increase: 0, decrease: 1 },
              { name: '测试', value: 7, increase: 0, decrease: 1 },
              { name: '数据', value: 8, increase: 0, decrease: 1 },
            ]
          ],
        },
        {
          uuid: '121231231231232',
          title: '项目情况',
          typeName: '柱状图',
          data: [
            [
              { name: '2020-1', value: 14 },
              { name: '2020-2', value: 20 },
              { name: '2020-3', value: 7 },
              { name: '2020-4', value: 8 },
            ]
          ],
        },
      ],
    };
    this.appService.fillSheet(excel, screen);
    const buffer = await excel.xlsx.writeBuffer();
    // const stream = new Readable();
    res.attachment('画板数据导出.xlsx');
    res.writeHead(200, {
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      // 'Content-Disposition': 'attachment; filename=画板数据导出.xlsx',
      'Content-Length': buffer.byteLength,
      'Cache-Control': 'no-cache, no-store, must-revalidate',
      Pragma: 'no-cache',
      Expires: 0,
    });
    // stream.push(buffer);
    // stream.pipe(res);
    res.end(buffer);
  }
}
