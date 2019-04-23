using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Xml;

/// <summary>
/// Summary description for DataTableToXml
/// </summary>
public class DataTableToXml
{
	public DataTableToXml()
	{
		//
		// TODO: Add constructor logic here
		//
	}

    static string ProcessDataTableToXml(DataTable dt, string NodeName, string RowsName)
    {
        string xmlstr = string.Empty;
        string tbsName = "DD" + Guid.NewGuid().ToString("N").Substring(0, 5);
        dt.TableName = tbsName;
        foreach (DataRow Rows in dt.Rows)
        {

            int CountRows = Rows.ItemArray.Length;
            int i = 0;
            for (i = 0; i < CountRows; i++)
            {

                if (string.IsNullOrEmpty(Rows[i].ToString()))
                {
                    try
                    {
                        Rows[i] = " ";
                    }
                    catch
                    {
                        if (dt.Columns[i].DataType == typeof(DateTime))
                            Rows[i] = DateTime.Parse("1900/01/01");
                        if (dt.Columns[i].DataType == typeof(int) ||
                            dt.Columns[i].DataType == typeof(Int16) ||
                            dt.Columns[i].DataType == typeof(Int32) ||
                             dt.Columns[i].DataType == typeof(Int64))
                            Rows[i] = 0;
                    }
                }

            }

        }

        System.IO.StringWriter writer = new System.IO.StringWriter();

        dt.WriteXml(writer, XmlWriteMode.WriteSchema, false);



        XmlDocument XmlDoc = new XmlDocument();
        XmlDoc.LoadXml(writer.ToString());
        XmlNodeList xnList = (XmlNodeList)XmlDoc.SelectNodes("//" + tbsName);

        foreach (XmlNode xn in xnList)
        {
            xmlstr += xn.OuterXml;
        }
        if (RowsName != null)
        {
            string x = @"<{0}>{1}</{0}>";
            x = string.Format(x, NodeName, xmlstr.Replace(tbsName, RowsName));
            return x;
        }
        else
        {
            string x = @"{0}";
            x = string.Format(x, xmlstr.Replace(tbsName, NodeName));
            return x;
        }
    }
    public static string ConvertDatatableToXML(DataTable dt, string NodeName, string RowsName)
    {
        return ProcessDataTableToXml(dt, NodeName, RowsName);
    }
    public static string ConvertDatatableToXML(DataTable dt, string NodeName)
    {
        return ProcessDataTableToXml(dt, NodeName, null);
    }
}