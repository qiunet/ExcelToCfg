import * as Path from "path";
import * as Excel from 'exceljs';
import ejs from "ejs";
import os from "os";
import fs from "fs";
import { ExcelSheet } from './ExcelSheet';

/**
 * 跟 common/Enums/Role 同步.
 * ExcelToCfg  是个单独能运行的ts.
 * 通用的枚举自己保存一份.
 */
export enum Role {
    /**
     * 客户端
     */
    CLIENT,
    /**
     * 服务器
     */
    SERVER,
    /**
     * 其它
     */
    OTHER,
}


export class Constants {
    /**
     * 配置目录
     */
    public static readonly SETTING_DIR = ".dTools";
    /**
     * 存放ejs的目录
     */
    public static readonly EJS_TEMPLATE_DIR = "ejs";

    /**
     * 获得ejs 模板路径
     */
    public static ejsTemplateDir(): string {
        return Path.join(os.homedir(), this.SETTING_DIR, this.EJS_TEMPLATE_DIR);
    }
}

/**
 * excel 转 cfg
 *
 * 输出类型的字符串给模板
 *
 * [
 *   [
 *     UnitData { fieldName: 'id', type: 'int', val: 1 },
 *     UnitData { fieldName: 'val', type: 'long', val: 200 }
 *   ],
 *   [
 *     UnitData { fieldName: 'id', type: 'int', val: 2 },
 *     UnitData { fieldName: 'val', type: 'long', val: 20 }
 *   ]
 * ]
 */
export class ExcelToCfg {
    /**
     * 角色类型
     * @private
     */
    private readonly _role: Role;
    /**
     * 相对于configDir路径
     * @private
     */
    private readonly _fileRelativePath: string;
    /**
     * 配置文件路径
     * @private
     */
    private readonly _configDir: string;
    /**
     * 输出文件夹路径
     *
     * @private
     */
    private readonly _outputDirPaths: string[];
    /**
     * 记录日志
     * @private
     */
    private readonly _logger: (info: string) => void;

    constructor(role: Role, fileRelativePath: string, configDir: string, outputDirPaths:string[], logger?: (info: string) => void) {
        this._fileRelativePath = fileRelativePath;
        this._outputDirPaths = outputDirPaths;
        this._configDir = configDir;
        this._role = role;
        if (logger) {
            this._logger = logger;
        }else {
            this._logger = console.log;
        }
    }

    /**
     * 获得文件名
     */
    getFileName():string {
        return Path.basename(this.fileRelativePath);
    }

    /**
     * 获得配置前缀
     */
    getCfgPrefix():string {
        let name = this.getFileName();
        return name.substring(name.indexOf("_") + 1, name.lastIndexOf("."));
    }


    get logger(): (info: string) => void {
        return this._logger;
    }

    get role(): Role {
        return this._role;
    }

    get fileRelativePath(): string {
        return this._fileRelativePath;
    }

    get configDir(): string {
        return this._configDir;
    }

    get outputDirPaths(): string[] {
        return this._outputDirPaths;
    }

    /**
     * 转换并生成文件
     */
    convert(): void {
        if (this._outputDirPaths === null || this._outputDirPaths.length === 0) {
            this.logger("项目输出路径为空. 转化终止!\n")
            return
        }

        if (! this._fileRelativePath.endsWith(".xlsx")) {
            this._logger(this._fileRelativePath + "不是xlsx文件");
            throw new Error(this._fileRelativePath + "不是xlsx文件");
        }
        let workbook = new Excel.Workbook();
        fs.promises.readFile(Path.join(this._configDir, this._fileRelativePath)).then(data => {
            workbook.xlsx.load(data.buffer).then(workbook => {
                for (let worksheet of workbook.worksheets) {
                    if (worksheet.name === 'end') {
                        break;
                    }

                    if (worksheet.name.startsWith("#")) {
                        continue;
                    }

                    new ExcelSheet(this, worksheet).handlerSheet();
                }
            });
        });
    }

    /**
     * 服务器脚本调用.
     * 转换整个文件夹.
     * @param configPath
     * @param outDirs
     * @param logger
     */
    public static convertDir(configPath: string, outDirs: string[], logger?: (info: string) => void) {
        this.roleConvert(Role.SERVER, configPath, '/', outDirs, logger);
    }


    public static roleConvert(role: Role, configPath: string, relativePath: string, outDirs: string[], logger?: (info: string) => void): void {
        let filePath = Path.join(configPath, relativePath);
        if (! fs.statSync(filePath).isDirectory()) {
            if (! relativePath.endsWith(".xlsx")) {
                return;
            }

            let excelToCfg = new ExcelToCfg(role, relativePath, configPath, outDirs, logger)
            excelToCfg.convert();
            return;
        }

        for (const file of fs.readdirSync(configPath)) {
            if (file.startsWith(".")) {
                // 一般 .svn .git 目录.
                continue;
            }

            this.roleConvert(role, configPath, Path.join(relativePath, file), outDirs, logger);
        }
    }
}
