import {Injectable} from '@angular/core';

const tableCols: any[] = [
  {fld: 'isp', hdr: 'ISP'},
  {fld: 'fnn', hdr: 'FNN Number'},
  {fld: 'pb', hdr: 'Peak Bandwidth (Kbps)'},
  {fld: 'nSio', hdr: 'Network SIO (count)'},
  {fld: 'kbps', hdr: 'Kbps/SIO'},
  {fld: 'curTier', hdr: 'Current Tier'},
  {fld: 'lastMonthTier', hdr: 'Last Month Tier'}
];

@Injectable()
export class TableService {
  tables: Row[][] = [];
  constructor() {
    for (let i = 0; i < 2; i++) {
      this.tables[i] = [];
      for (let j = 0; j < 9; j++) {
        let row: Row = new Row();
        row.fnn = 'N' + (Math.floor(Math.random()*900000)+100000) + 'K';
        row.pb = Number(((Math.random() * 100) + 50).toFixed(3));
        row.nSio = Math.floor(Math.random()*9000) + 1000;
        row.kbps = Math.floor(Math.random()*4500) + 500;
        row.curTier = Math.floor(Math.random()*20) + 1;
        row.lastMonthTier = Math.floor(Math.random()*20) + 1;
        this.tables[i][j] = row;
      }
    }
    for (let i = 0; i < 2; i++) {
      this.tables[i][0].isp = 'AAPT';
      this.tables[i][1].isp = 'Amoore';
      this.tables[i][2].isp = 'Brennan';
      this.tables[i][3].isp = 'Ettel';
      this.tables[i][4].isp = 'IInet';
      this.tables[i][5].isp = 'Issek';
      this.tables[i][6].isp = 'Optus';
      this.tables[i][7].isp = 'People Telecom';
      this.tables[i][8].isp = 'XYZnet';
    }
  }
  getTableCols(): any[] {
    return tableCols;
  }
  getTables() {
    return this.tables;
  }
}
export class Row {
  constructor(public isp?, public fnn?, public pb?, public nSio?, public kbps?, public curTier?, public lastMonthTier?) {}
}

