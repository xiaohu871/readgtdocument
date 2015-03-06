using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Specialized;
using System.Reflection;
using Files;
namespace ReadGTDocument
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IniFiles iniconfig = new IniFiles(AppDomain.CurrentDomain.BaseDirectory + "\\config.ini");
            ExcelTranslater et = new ExcelTranslater();
            ExcelRangeClass range = new ExcelRangeClass();
            range.Rowcount = 12;
            range.Colcount = 9;
            range.F2InCount = 4;

            ExcelRangeClass rangeufx = new ExcelRangeClass();
            rangeufx.Rowcount = 14;
            rangeufx.Colcount = 19;
            rangeufx.F2InCount = 3;
            rangeufx.Startrow = 3;
            rangeufx.Startcol = 2;
            rangeufx.Points = new Dictionary<string, Point>();
            rangeufx.Propname_map = new NameValueCollection();
            iniconfig.ReadSectionValues("PropName2ExcelItemName" ,rangeufx.Propname_map);

            FunctionClass [] fs = new FunctionClass[2]{ new FunctionClass(),new FunctionClass()};
            fs[0].Function_id = "200";
            fs[0].In_fields  = new FieldClass[3]{new FieldClass(),new FieldClass(),new FieldClass()};
            fs[0].Out_fields = new FieldClass[4]{new FieldClass(), new FieldClass(), new FieldClass(), new FieldClass()};
            fs[0].In_fields[0].Name = "字段1";
            fs[0].In_fields[1].Name = "字段2";
            fs[0].In_fields[2].Name = "字段3";
            fs[0].Out_fields[0].Name = "o字段1";
            fs[0].Out_fields[1].Name = "o字段2";
            fs[0].Out_fields[2].Name = "o字段3";
            fs[0].Out_fields[3].Name = "o字段4";
            fs[0].Is_resultset = "xxx";

            fs[1].Function_id = "201";
            fs[1].Function_name = "修改密码";
            fs[1].In_fields = new FieldClass[3] { new FieldClass(), new FieldClass(), new FieldClass() };
            fs[1].Out_fields = new FieldClass[4] { new FieldClass(), new FieldClass(), new FieldClass(), new FieldClass() };
            fs[1].In_fields[0].Name = "字段1";
            fs[1].In_fields[1].Name = "字段2";
            fs[1].In_fields[2].Name = "字段3";
            fs[1].Out_fields[0].Name = "o字段1";
            fs[1].Out_fields[1].Name = "o字段2";
            fs[1].Out_fields[2].Name = "o字段3";
            fs[1].Out_fields[3].Name = "o字段4";

            //et.WriteItnFile(fs, AppDomain.CurrentDomain.BaseDirectory + "投资赢家2.0期货交易接口规范-ITN.xls", "功能接口", range);
            et.WriteItnFile(fs, AppDomain.CurrentDomain.BaseDirectory + "投资赢家2.0期货交易接口规范-ITN.xls", "UFX", rangeufx);
        }
    }
}
