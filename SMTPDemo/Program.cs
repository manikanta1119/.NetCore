using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Newtonsoft.Json;
using System.IO;

namespace SMTPDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            //SmtpClient client = new SmtpClient();
            //client.Host = "smtp.gmail.com";
            //client.Port = 587;
            //client.EnableSsl = true;
            
            
            //client.DeliveryMethod = SmtpDeliveryMethod.Network;
            //client.UseDefaultCredentials = false;
            
            //client.Credentials=new NetworkCredential("manip0811@gmail.com", "premchand9697");
            
            //MailMessage mail = new MailMessage();
            //mail.To.Add("pmanikantachowdary@gmail.com");
            //mail.From = new MailAddress("manip0811@gmail.com");
            //mail.Subject = "Test";
            //mail.Body = "This is Testing";
            //client.Send(mail);
            //Console.WriteLine("Sucessfully send");
            //Console.ReadKey();
            GetDataTable();
            

        }

        public static DataTable GetDataTable()
        {
            DataTable dt = new DataTable();
            SqlConnection connection = new SqlConnection("Data Source=US5CD9011ZYR\\MSSQLMANI;Initial Catalog=Demo;Integrated Security=True");
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandText= "SELECT *FROM[Demo].[sales].[customers]";
            if(connection.State !=ConnectionState.Open)
            {
                connection.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
               
            }
            string path = "C:/Users/mparupally/Documents/Visual Studio 2017/Workspace/Export.csv";
            Converttosv(dt, path);
            return dt;
            
        }

        public static void Converttosv(DataTable datatable,string path)
        {
            StreamWriter sw = new StreamWriter(path,false);

            for(int i=0;i<datatable.Columns.Count;i++)
            {
                sw.Write(datatable.Columns[i]);
                if(i<datatable.Columns.Count -1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach(DataRow dr in datatable.Rows)
            {
                for(int i=0;i<datatable.Columns.Count;i++)
                {
                    if(!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();

                        if(value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < datatable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
                   
                
     

        
        ////public static string ConvertDatatabletoJson(DataTable datatable)
        //{
        //    JsonSerializer serializer = new JsonSerializer();

        //    List<Dictionary<string, object>> tableRows = new List<Dictionary<string, object>>();

        //    Dictionary<string, object> row;

        //    foreach (DataRow dr in datatable.Rows)
        //    {
        //        row = new Dictionary<string, object>();
        //        foreach (DataColumn col in datatable.Columns)
        //        {
        //            row.Add(col.ColumnName, dr[0]);
        //        }
        //        tableRows.Add(row);
        //    }
        //    string result = JsonConvert.SerializeObject(tableRows);
        //    Console.WriteLine(result);

        //    Console.Read();
        //    TransfertoEXCEL(result);
        //    return result;

            

        //}

        ////public static string TransfertoEXCEL(string data)
        //{

        //    var datatable = (DataTable)JsonConvert.DeserializeObject(data, (typeof(DataTable)));

        //    var lines = new List<string>();

        //    string[] columnNames = datatable.Columns.Cast<DataColumn>().Select(Colum => Colum.ColumnName).ToArray();
        //    var header = string.Join(",", columnNames);

        //    lines.Add(header);

            
            
        //    StringBuilder sb = new StringBuilder();
        //    foreach (DataRow row in datatable.Rows)
        //    {
        //        IEnumerable<string> feilds = row.ItemArray.Select(feild => feild.ToString());
        //         sb.AppendLine(string.Join(",", feilds));
                
        //    }
        //    //var valueLines = datatable.AsEnumerable()
        //    //.Select(row => string.Join(",", row.ItemArray));
        //    //StringBuilder sb = new StringBuilder();
        //    //foreach (var datarow in datatable.Rows)
        //    //{
        //    //    for(int i=0; i<datatable.Columns.Count; i++)
        //    //    {
        //    //        sb.Append(datarow[i].ToString() + ",");
        //    //    }
        //    //}

        //    string[] stringarray = sb.ToString().Split(' ').ToArray();


        //    File.WriteAllLines(@"C:/Users/mparupally/Documents/Visual Studio 2017/Workspace/Export.csv", stringarray);


        //    return (data);
        //}
    }

    
}
