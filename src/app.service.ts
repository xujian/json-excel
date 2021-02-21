import { Injectable } from '@nestjs/common';
import Excel = require('exceljs');

export interface Component {
  uuid: string,
  title: string;
  typeName: string;
  data: any[][];
}

export type Cell = {
  col: number,
  letter: string,
  row: number,
};

const alphabet: string[] = Array(26)
  .fill('')
  .map((_a: any, i: number) => String.fromCharCode(i + 65),
);

/**
 * 解析单元格记号
 * 将C3解析为{col: 1, letter: 'C', row: 3}格式
 * @param cell 单元格记号
 * @returns
 */
function parseCell(cell: string): Cell {
  const [, letter, row] = cell.match(/^([A-Z]+)(\d+)$/),
    col = alphabet.indexOf(letter) + 1;
  return {
    col,
    letter,
    row: +row,
  };
}

/**
 * MaxView工程导出的页面数据
 */
export interface Screen {
  title: string;
  components: Component[];
}

@Injectable()
export class AppService {
  getHello(): string {
    return 'Hello World!';
  }

  buildExcel(): Excel.Workbook {
    const book = new Excel.Workbook();
    book.model.title = 'MaxView数据导出';
    book.creator = 'MaxView';
    book.lastModifiedBy = 'MaxView';
    const date = new Date();
    book.created = date;
    book.modified = date;
    book.lastPrinted = date;
    book.properties.date1904 = true;
    book.calcProperties.fullCalcOnLoad = true;
    book.views = [
      {
        x: 0,
        y: 0,
        width: 1600,
        height: 800,
        firstSheet: 0,
        activeTab: 1,
        visibility: 'visible',
      },
    ];
    return book;
  }

  async readExcel(): Promise<Excel.Workbook> {
    const book = new Excel.Workbook();
    await book.xlsx.readFile('./output-template.xlsx');
    return book;
  }

  private printComponent(
    sheet: Excel.Worksheet,
    column: string,
    row: number,
    component: Component,
  ): number {
    sheet.getCell(`${column}${row++}`).value = component.title;
    sheet.getCell(`${column}${row++}`).value = component.typeName;
    for (const d of component.data) {
      row = this.printData(sheet, column, row, d);
    }
    return row;
    // sheet.commit();
  }

  private printData(
    sheet: Excel.Worksheet,
    column: string,
    row: number,
    data: any[]): number {
    const letters: string[] = [...Array(26)].map((n, i) =>
      String.fromCharCode(i + 65),
    );
    let x = letters.indexOf(column),
      y = row;
    for (const d of data) { // loop data rows
      x = letters.indexOf(column);
      if (y === row) {
        for (const k in d) { // print header
          sheet.getCell(`${letters[x++]}${y}`).value = k;
        }
        x = letters.indexOf(column);
      }
      for (const k in d) { // loop date keys
        sheet.getCell(`${letters[x++]}${y}`).value = d[k];
      }
      y++;
    }
    y = y + 2;
    return y;
  }

  /**
   * 将组件标准数据转换为table rows
   * @param component 
   * @returns 
   */
  buildTableRows(component: Component) {
    const rows = component.data[0].map((d) => {
      const { name, value, ...rest } = d,
        keys = ['name', 'value', ...Object.keys(rest)];
      return keys.map((k) => d[k]);
    });
    const d = component.data[0][0];
    const { name, value, ...rest } = d,
      keys = ['name', 'value', ...Object.keys(rest)],
      columns = keys.map(k => ({
        name: `(${k})`,
        filterButton: false,
      }))
    return { columns, rows };
  }

  /**
   * 获取表格占据的范围
   * @param table 表格
   * @returns 
   */
  getSpread(table: Excel.TableProperties): [Cell, Cell] {
    const columns = table.columns.length,
      rows = table.rows.length;
    const from = parseCell(table.ref),
      right = from.col + columns - 1,
      letter = alphabet[right - 1],
      bottom = from.row + rows; // 包含table header
    return [
      from,
      {
        col: right,
        letter,
        row: bottom,
      },
    ];
  }

  /**
   * 格式化表格
   * 设置表格样式
   * @param sheet 
   * @param table 
   */
  formatTable(sheet: Excel.Worksheet, table: Excel.TableProperties): void {
  }

  /**
   * 将页面组件数据写入worksheet
   * @param book 
   * @param screen 
   */
  fillSheet(book: Excel.Workbook, screen: Screen): void {
    const sheet = book.addWorksheet(screen.title || '页面', {
      views: [{ showGridLines: false }],
    });
    const firstColumn = sheet.getColumn('A');
    firstColumn.width = 2;
    let anchor = 'B3';
    screen.components.forEach(c => {
      const { columns, rows } = this.buildTableRows(c)
      const table: Excel.TableProperties = {
        name: `Component${c.uuid.replace('-', '')}`,
        ref: anchor,
        headerRow: true,
        totalsRow: false,
        style: {
          theme: 'TableStyleDark3',
          showRowStripes: true,
        },
        columns,
        rows,
      };
      sheet.addTable(table);
      const [from, to] = this.getSpread(table);
      console.log('formatTable.................', from, to, anchor);
      anchor = `${from.letter}${to.row + 2}`;
      this.formatTable(sheet, table);
    });
  }
}
