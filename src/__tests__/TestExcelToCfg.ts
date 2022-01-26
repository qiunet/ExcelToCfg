import { ExcelToCfg, Role } from '../index';

let excelToCfg = new ExcelToCfg(Role.SERVER, "G全局表_global_setting.xlsx", __dirname, [__dirname]);
excelToCfg.convert()
