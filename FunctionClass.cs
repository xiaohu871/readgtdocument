using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Collections.Specialized;

namespace ReadGTDocument
{   
    enum DataType : byte　　//显示指定枚举的底层数据类型
    { 
        dtString,
        dtInt,　　
        dtDouble
    };
    enum RequiredType : byte
    {
        rtRequired,
        rtConditional,
        rtOptional
    }
    class FieldClass
    {
        private string _name; //名称

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        private DataType _data_type; //数据类型

        public DataType Data_type
        {
            get { return _data_type; }
            set { _data_type = value; }
        }
        private string _len; //长度

        public string Len
        {
            get { return _len; }
            set { _len = value; }
        }
        private string _desc; //说明

        public string Desc
        {
            get { return _desc; }
            set { _desc = value; }
        }
        private RequiredType _required_type; //出现要求

        public RequiredType Required_type
        {
            get { return _required_type; }
            set { _required_type = value; }
        }
        private string _default_value; //默认值

        public string Default_value
        {
            get { return _default_value; }
            set { _default_value = value; }
        }

        private string _remark; //备注

        public string Remark
        {
            get { return _remark; }
            set { _remark = value; }
        }
    };
    class ErrorClass
    {
        private string _error_no;

        public string Error_no
        {
            get { return _error_no; }
            set { _error_no = value; }
        }
        private string _error_info;

        public string Error_info
        {
            get { return _error_info; }
            set { _error_info = value; }
        }
    }
    class UpdatedClass
    {
        private string _updated_date; //更新日期

        public string Updated_date
        {
            get { return _updated_date; }
            set { _updated_date = value; }
        }
        private string _updated_content; //更新内容

        public string Updated_content
        {
            get { return _updated_content; }
            set { _updated_content = value; }
        }
        private string _updated_person;//修改人

        public string Updated_person
        {
            get { return _updated_person; }
            set { _updated_person = value; }
        }
    }
    class FunctionClass
    {
        private string _function_id; //功能号

        public string Function_id
        {
            get { return _function_id; }
            set { _function_id = value; }
        }
        private string _Sfunction_id; //短功能号

        public string SFunction_id
        {
            get {
                if ((_Sfunction_id == null ) || (_Sfunction_id.Trim() == ""))
                    return _function_id;
                else
                    return _Sfunction_id;
            }
            set { _Sfunction_id = value; }
        }     

        string _region;
        public string Region
        {
            get { return _region; }
            set { _region = value; }
        }  
        private string _function_name;//功能名称

        public string Function_name
        {
            get { return _function_name; }
            set { _function_name = value; }
        }
        private string _function_desc; //功能说明

        public string Function_desc
        {
            get { return _function_desc; }
            set { _function_desc = value; }
        }
        private string _is_resultset; //是否结果集

        public string Is_resultset
        {
            get { return _is_resultset; }
            set { _is_resultset = value; }
        }
        private string _last_updated; //最后更新日期

        public string Last_updated
        {
            get { return _last_updated; }
            set { _last_updated = value; }
        }
        private string _version; //版本号

        public string Version
        {
            get { return _version; }
            set { _version = value; }
        }
        private FieldClass[] _in_fields; //输入参数

        internal FieldClass[] In_fields
        {
            get { return _in_fields; }
            set { _in_fields = value; }
        }
        private FieldClass[] _out_fields; //输出参数

        internal FieldClass[] Out_fields
        {
            get { return _out_fields; }
            set { _out_fields = value; }
        }
        private string _remark; //业务说明

        public string Remark
        {
            get { return _remark; }
            set { _remark = value; }
        }
        private ErrorClass[] _Errors; //出错说明

        internal ErrorClass[] Errors
        {
            get { return _Errors; }
            set { _Errors = value; }
        }
        private UpdatedClass[] _UpdatedDesc;//更新说明

        internal UpdatedClass[] UpdatedDesc
        {
            get { return _UpdatedDesc; }
            set { _UpdatedDesc = value; }
        }
    };
    class ExcelRangeClass
    {
        Dictionary<string, Point> _Points;

        public Dictionary<string, Point> Points
        {
            get { return _Points; }
            set { _Points = value; }
        }

        private NameValueCollection _propname_map;

        public NameValueCollection Propname_map
        {
            get { return _propname_map; }
            set { _propname_map = value; }
        }
        int _startrow;
        
        public int Startrow
        {
            get {
                if (_startrow <= 0)
                    return 2;
                else return _startrow; 
                }
            set { _startrow = value; }
        }
        int _startcol;

        public int Startcol
        {
            get {
                if (_startcol <= 0)
                    return 2;
                else 
                    return _startcol; 
                }
            set { _startcol = value; }
        }
        int _rowcount;

        public int Rowcount
        {
            get { return _rowcount; }
            set { _rowcount = value; }
        }
        int _colcount;

        public int Colcount
        {
            get { return _colcount; }
            set { _colcount = value; }
        }
        int _f2InCount;

        public int F2InCount
        {
            get { return _f2InCount; }
            set { _f2InCount = value; }
        }
        //public ExcelRangeClass(Dictionary<string, string> propname_map, Dictionary<string, Point> points)
        //{
        //    _propname_map = propname_map;
        //    _Points = points;
        //}
    }
}
