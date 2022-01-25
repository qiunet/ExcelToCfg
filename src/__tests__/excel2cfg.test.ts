import { ExcelToCfg, Role } from '../index';

test("convert", () => {
    let excelToCfg = new ExcelToCfg(Role.SERVER, "G全局表_global_setting.xlsx", __dirname, __dirname);
    excelToCfg.convert()
});
