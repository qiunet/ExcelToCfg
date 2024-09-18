import * as Excel from 'exceljs';
import { Utils } from './utils/Utils';
import ejs from 'ejs';
import Path from 'path';
import fs from 'fs';
import { Constants, ExcelToCfg, Role } from './index';

/**
 * 列输出类型
 */
enum OutputType {
    /**
     * 忽略该列
     */
    IGNORE,
    /**
     * 仅客户端
     */
    CLIENT,
    /**
     * 仅服务器
     */
    SERVER,
    /**
     * 所有角色
     */
    ALL,

}

/**
 * 所有的excel配置数据
 */
class ExcelCfg {
    /**
     * 配置数据 行数组
     */
    private rows: ExcelRowCfg[] = [];

    /**
     * 添加数据
     * @param rowData
     */
    public addData(rowData: ExcelRowCfg): void{
        this.rows.push(rowData);
    }
}

/**
 * 行数据
 */
class ExcelRowCfg {
    cells: UnitData[] = [];

    /**
     * 添加数据
     * @param data
     * @private
     */
    public addData(data: UnitData): void {
        this.cells.push(data);
    }
}
/**
 * 单元的数据
 */
class UnitData {
    /**
     * 字段名称
     */
    name: string = "";
    /**
     * 类型描述
     * 一般就int long
     * string
     * (int[]  long[]是string 类型, 用;隔开)
     */
    type: string = "string";
    /**
     * 字段值
     */
    val: string|number = '';

    /**
     * string type
     */
    isStringType():boolean {
        return ! this.isNumberType();
    }

    /**
     * number 类型
     */
    isNumberType() {
        return  this.type === 'int' || this.type === 'long';
    }
}

/**
 * 正确字符串表述值
 * @type {Set<string>}
 */
const trueValSet = new Set(["yes", "1", "true"]);

export class ExcelSheet {
    /**
     * ExcelToCfg  的配置
     * @private
     */
    private readonly cfgConfig: ExcelToCfg;
    /**
     * 单个sheet
     * @private
     */
    private readonly sheet: Excel.Worksheet;
    /**
     * 日志
     * @private
     */
    private readonly logger: (info: string) => void;
    /**
     * 真实列数
     * @private
     */
    private readonly columnCount: number;
    /**
     * 真实行数
     * @private
     */
    private readonly rowCount: number;
    /**
     * 字段名
     * @private
     */
    private readonly fieldNames: string[];
    /**
     * 字段类型
     * @private
     */
    private readonly fieldTypes: string[];
    /**
     * 输出类型
     * @private
     */
    private readonly fieldOutputTypes:OutputType[];

    constructor(cfgConfig: ExcelToCfg, sheet: Excel.Worksheet) {
        this.logger = cfgConfig.logger;
        this.cfgConfig = cfgConfig;
        this.sheet = sheet;

        this.columnCount = this.actualColumnLength();
        this.rowCount = this.actualRowLength();
        this.fieldOutputTypes = this.getFieldOutputTypes();
        this.fieldNames = this.getFieldNames();
        this.fieldTypes = this.getFieldTypes();

        this.logger("处理excel["+cfgConfig.getFileName()+"] ["+cfgConfig.getCfgPrefix()+"_"+ sheet.name + "] 行数:"+this.rowCount+" 列数:"+this.columnCount + "\n");
    }

    /**
     * 得到cell的值.
     * @param cell
     * @param type 类型
     */
    private cellValue(cell:Excel.Cell, type : string = 'string'): number|string {
        function getterVal(){
            if (cell.value == null) {
                return '';
            }

            if (cell.type == Excel.ValueType.Hyperlink) {
                return cell.text
            }

            if (cell.type == Excel.ValueType.Formula) {
                return (<Excel.CellFormulaValue>cell.value).result + '';
            }

            if (cell.type == Excel.ValueType.RichText) {
                let result = '';
                let richText: Excel.RichText [] = (<Excel.CellRichTextValue>cell.value).richText;
                richText.forEach(r => result += r.text);
                return result;
            }

            if (cell.type == Excel.ValueType.Date) {
                return Utils.dateFormat(<Date>cell.value);
            }

            return cell.value.toString().trim();
        }
        let result = getterVal();
        if (type === 'long' || type === 'int') {
            return Number(result);
        }
        return result.replaceAll("\"", "\\\"");
    }
    /**
     * 实际的列数
     * @param sheet
     */
    private actualColumnLength() {
        let row = this.sheet.getRow(4);
        let index = 1;
        for (; index <= row.actualCellCount;) {
            if (this.cellValue(row.getCell(index)) === '') {
                return index - 1;
            }
            index++;
        }
        return Math.min(row.actualCellCount, index);
    }

    /**
     * 实际的行数
     * @param sheet
     */
    private actualRowLength() {
        let index = 0;
        for (; index <= this.sheet.actualRowCount;) {
            if (this.cellValue(this.sheet.getRow(index + 1).getCell(1)) === '') {
                break;
            }
            index++;
        }
        return Math.min(this.sheet.actualRowCount, index);
    }

    /**
     * 获取fieldNames
     * @private
     */
    private getFieldNames():string[] {
        let fieldNames: string[] = [];
        let nameRow = this.sheet.getRow(2)
        for (let i = 1; i <= this.columnCount; i++) {
            fieldNames.push(this.cellValue(nameRow.getCell(i)).toString());
        }
        return fieldNames;
    }
    /**
     * 获取fieldTypes
     * @private
     */
    private getFieldTypes():string[] {
        let fieldTypes: string[] = [];
        let typeRow = this.sheet.getRow(3)
        for (let i = 1; i <= this.columnCount; i++) {
            fieldTypes.push(this.cellValue(typeRow.getCell(i)).toString());
        }
        return fieldTypes;
    }

    /**
     * 输出类型
     * @private
     */
    private getFieldOutputTypes(): OutputType[] {
        let fieldOutputTypes:OutputType[] = [];
        let createTypeRow = this.sheet.getRow(4)
        for (let i = 1; i <= this.columnCount; i++) {
            let outputTypeString:string = this.cellValue(createTypeRow.getCell(i)).toString();
            let outputType:OutputType = (<any>OutputType)[outputTypeString];
            fieldOutputTypes.push(outputType)
        }
        return fieldOutputTypes;
    }

    /**
     * 真实名称
     * @private
     */
    private realName():string {
        let splits = this.sheet.name.split(".");
        return splits[splits.length - 1];
    }

    /**
     * 输出的模板类型
     * @private
     */
    private outputTemplates(): Set<string> {
        let set = new Set(this.sheet.name.split("."));
        set.delete(this.realName());
        let justClient = set.has("c");
        let justServer = set.has("s");
        set.delete("c");
        set.delete("s");

        if (justServer) {
            // 服务器只用json格式
            set.clear();
        }

        if (! justClient || set.size === 0) {
            set.add('json');
        }
        return set;
    }

    /**
     * 判断全表是否都没有该角色需要的字段
     */
    private handlerNonFieldRoleNeed(outputType: OutputType): boolean {
        var nonFieldNeed = true;
        for (let type of this.fieldOutputTypes) {
            if (type == OutputType.ALL || type == outputType) {
                nonFieldNeed = false;
                break
            }
        }
        return nonFieldNeed;
    }

    /**
     * 处理单个sheet
     * @param sheet sheet
     * @private
     */
    public handlerSheet() {
        if (this.columnCount <= 0) {
            this.logger("Excel["+this.cfgConfig.getFileName()+"] 列数["+this.columnCount+"]错误, 没有任何数据!")
            return;
        }
        if (this.rowCount < 0) {
            this.logger("Excel["+this.cfgConfig.getFileName()+"] 行数["+this.rowCount+"]错误, 没有表头!")
            return;
        }

        if (this.cfgConfig.role === Role.SERVER
        && (this.sheet.name.indexOf("c.") !== -1 || this.handlerNonFieldRoleNeed(OutputType.SERVER))) {
            this.logger('[' + this.cfgConfig.getFileName() + '.' + this.sheet.name + ']没有字段符合[SERVER]角色, 输出忽略!')
            return;
        }


        if (this.cfgConfig.role === Role.CLIENT
        && (this.sheet.name.indexOf("s.") !== -1 || this.handlerNonFieldRoleNeed(OutputType.CLIENT))) {
            this.logger('[' + this.cfgConfig.getFileName() + '.' + this.sheet.name + ']没有字段符合[CLIENT]角色, 输出忽略!')
            return;
        }

        let ignoreFieldIndex = -1;
        this.fieldNames.find((element, index, arr) => {
            if (element === '.ignore') {
                ignoreFieldIndex = index;
            }
        });

        let excelCfg: ExcelCfg = new ExcelCfg();
        for (let i = 5; i <= this.rowCount; i++) {
            let row = this.sheet.getRow(i);
            if (ignoreFieldIndex >= 0
                && trueValSet.has(this.cellValue(row.getCell(ignoreFieldIndex + 1)).toString().toLowerCase()))
            {
                continue;
            }

            let dataList: ExcelRowCfg = new ExcelRowCfg();
            for (let j = 1; j <= this.columnCount; j++) {
                let outputType = this.fieldOutputTypes[j - 1];
                if (outputType === OutputType.IGNORE || j === (ignoreFieldIndex + 1)) {
                    continue;
                }

                if (this.cfgConfig.role === Role.CLIENT && outputType === OutputType.SERVER) {
                    continue;
                }

                if (this.cfgConfig.role === Role.SERVER && outputType === OutputType.CLIENT) {
                    continue;
                }

                let data: UnitData = new UnitData();
                data.type = this.fieldTypes[j - 1];
                data.name = this.fieldNames[j - 1];
                data.val = this.cellValue(row.getCell(j), data.type);
                dataList.addData(data);
            }
            excelCfg.addData(dataList);
        }

        this.output(excelCfg, this.outputTemplates());
    }
    private readonly jsonEjs: string = "[\n" +
        "<%_ rows.forEach(function (row, rIndex){ _%>\n" +
        "    {\n" +
        "    <%_ row.cells.forEach(function (cell, cIndex) { _%>\n" +
        "        \"<%= cell.name %>\": <% if (cell.isStringType()) {%>\"<% }%><%-cell.val%><% if (cell.isStringType()) {%>\"<% }%><% if(cIndex < row.cells.length - 1) { %>,<% } %>\n" +
        "     <%_}); _%>\n" +
        "    }<% if(rIndex < rows.length - 1) { %>,<% } %>\n" +
        "<%_ }); _%>\n" +
        "]\n";
    /**
     * 处理数据
     * @param data
     * @param outputPath
     * @param templateDir
     * @param useTypes
     */
    public output(data: ExcelCfg, useTypes: Set<string>) {
        useTypes.forEach(postfix => {
            if (postfix === 'json') {
                return this.output0(ejs.render(this.jsonEjs, data), postfix);
            }
            const templateFilePath:string = Path.join(Constants.ejsTemplateDir(), postfix + ".ejs");
            ejs.renderFile(templateFilePath, data, ((err, content) => {
                if (err) {
                    console.error(err);
                    return;
                }
                this.output0(content, postfix);
            }));
        });
    }

    /**
     * 往文件输出内容
     * @param content
     * @param postfix
     * @private
     */
    private output0(content: string, postfix: string): void {
        let fileName: string =  this.cfgConfig.getCfgPrefix() + "_" + this.realName() + "."+postfix;
        this.cfgConfig.outputDirPaths.forEach((outputDir) => {
            let filePath = Path.join(outputDir, Path.dirname(this.cfgConfig.fileRelativePath), fileName);
            if (!fs.existsSync(Path.dirname(filePath))) {
                fs.mkdirSync(Path.dirname(filePath), { recursive: true });
            }
            fs.writeFile(filePath, content, (err) => {
                if (err) {
                    console.error(err);
                }
            });
        });
    }
}
