export class Utils {
    /**
     * 对数字进行固定长度输出字符串 不够前面补0
     * @param val
     * @param length
     */
    public static fixedNumber(val: number, length: number) : string{
        return (Array(length).join("0") + val).slice(-length);
    }

    /**
     * 格式化日期为 yyyy-MM-dd HH:mm:ss
     * @param date
     */
    public static dateFormat(date: Date): string {
        return this.fixedNumber(date.getFullYear(), 4)+"-"+this.fixedNumber((date.getMonth() + 1), 2)+"-"+this.fixedNumber(date.getDate(), 2) +
            " "+this.fixedNumber(date.getHours(), 2)+":"+this.fixedNumber(date.getMinutes(), 2)+":"+this.fixedNumber(date.getSeconds(), 2);
    }
}
