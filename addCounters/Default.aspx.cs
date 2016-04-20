using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Npgsql;
using System.Text;
using System.Diagnostics;

namespace addCounters
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        protected void Button3_Click(object sender, EventArgs e)
        {
            String savePath = @"c:\temp\";
            if (FileUpload1.HasFile && FileUpload1.FileName.Substring(FileUpload1.FileName.Length - 4).ToLower() == "xlsx")
            {
                StreamWriter success = new StreamWriter(savePath + @"success.txt", false, Encoding.Default);
                StreamWriter errorRow = new StreamWriter(savePath + @"error.txt", false, Encoding.Default);
                String fileName = FileUpload1.FileName;
                //savePath += fileName;
                FileUpload1.SaveAs(savePath + fileName);
                var wb = new ClosedXML.Excel.XLWorkbook(savePath + fileName);
                string dat_uchet = Convert.ToString(wb.Worksheet(1).Row(3).Cell(6).Value).Trim();
                string date_pay = Convert.ToString(wb.Worksheet(1).Row(3).Cell(7).Value).Trim();
                for (int i = 5; i <= 5; i++)
                {
                    if (Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value).Trim() != "")
                    {
                        if (wb.Worksheet(1).Row(i).Cell(6).Value != null &&
                            Convert.ToString(wb.Worksheet(1).Row(i).Cell(6).Value).Trim() != "" &&
                            wb.Worksheet(1).Row(i).Cell(7).Value != null &&
                            Convert.ToString(wb.Worksheet(1).Row(i).Cell(7).Value).Trim() != "")
                        {

                            List<string> nzp_kvar =
                                SelectNzpKvar(Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                                    Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value).Trim());
                            if (nzp_kvar != null)
                            {

                                int nzp_serv = SelectNzpServ(Convert.ToString(wb.Worksheet(1).Row(i).Cell(5).Value).Trim());

                                string addCounters = AddCounter(
                                    nzp_kvar[0],
                                    nzp_kvar[1],
                                    Convert.ToString(wb.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                    dat_uchet,
                                    date_pay,
                                    Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(6).Value),
                                    Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(7).Value),
                                    nzp_serv
                                    );
                                if (addCounters == "Success")
                                {
                                    success.WriteLine("SUCCESS||" +
                                                      Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value).Trim() +
                                                      "||" + Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value).Trim() +
                                                      "||" + Convert.ToString(wb.Worksheet(1).Row(i).Cell(3).Value).Trim());
                                }
                                else
                                {
                                    errorRow.WriteLine(addCounters + "||" +
                                                       Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value).Trim() +
                                                       "||" + Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value).Trim() +
                                                       "||" + Convert.ToString(wb.Worksheet(1).Row(i).Cell(3).Value).Trim());
                                }
                            }
                            else
                            {
                                errorRow.WriteLine("Не удалось одназначно определить квартиру||" +
                                                   Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value).Trim() +
                                                   "||" + Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value).Trim() +
                                                   "||" + Convert.ToString(wb.Worksheet(1).Row(i).Cell(3).Value).Trim());
                            }
                        }
                    }
                    else
                    {
                        break;
                    }

                }
                
                success.Close();
                errorRow.Close();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                string targetCompressName = @"C:\Temp\info.zip";
                string filetozip = null;
                filetozip = "\"" + savePath + @"success.txt" + " " + "\"" + savePath + @"error.txt" + " ";
                startInfo.Arguments = "a -tzip \"" + targetCompressName + "\" \"" + filetozip + "\" -mx=9";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process x = Process.Start(startInfo);
                x.WaitForExit();

                string path = targetCompressName;
                byte[] bts = System.IO.File.ReadAllBytes(path);
                System.IO.File.Delete(targetCompressName);
                System.IO.File.Delete(savePath + @"success.txt");
                System.IO.File.Delete(savePath + @"error.txt");
                System.IO.File.Delete(savePath + fileName);
                Response.Clear();
                Response.ClearHeaders();
                Response.AddHeader("Content-Type", "Application/octet-stream");
                Response.AddHeader("Content-Length", bts.Length.ToString());
                Response.AddHeader("Content-Disposition", "attachment; filename=info.zip");
                Response.BinaryWrite(bts);
                Response.Flush();
                Response.End();
                System.IO.File.Delete(targetCompressName);
                Label1.Text = "Файл создан";
            }
            else
            {
                // Notify the user that a file was not uploaded.
                Label1.Text = "Необходимо загрузить фаил в формате .xlsx";
            }
        }

        public string SelectPkod(string address, string pkodPart, string fio)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText = @"SELECT pkod FROM fbill_data.kvar k
inner join fbill_data.dom d on k.nzp_dom = d.nzp_dom
inner join fbill_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
where pkod || '' like '" + pkodPart + "%' AND upper(ul.ulica) = upper('" + address.Split(',')[0] + "') AND d.ndom = '" + address.Split(',')[1].Split('-')[0]
                         + "' AND k.nkvar = '" + address.Split(',')[1].Split('-')[1] + "' AND replace(upper(fio), ' ','') = replace(upper('" + fio + "'), ' ','')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return "00";
                }
            }
            catch (Exception e)
            {
                return "0";
            }
        }

        public List<string> SelectNzpKvar(string address, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billGeu1;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText = @"SELECT nzp_kvar, num_ls 
FROM fbill_data.kvar k
inner join fbill_data.dom d on k.nzp_dom = d.nzp_dom
inner join fbill_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
where replace(upper(ul.ulica || ' д.' || d.ndom), ' ','') = replace(upper('" + address + "'), ' ','') AND 'кв. ' || k.nkvar || ' комн. ' ||  k.nkvar_n = '" + nkvar + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    return new List<string>()
                    {
                        dt.Rows[0][0].ToString(), 
                        dt.Rows[0][1].ToString()
                    };
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public int SelectNzpServ(string service)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billGeu1;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText = @"SELECT nzp_serv FROM fbill_kernel.services where service = '" + service + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    return Convert.ToInt32(dt.Rows[0][0].ToString());

                }
                else
                {
                    return 0;
                }
            }
            catch (Exception e)
            {
                return 0;
            }
        }

        public String AddCounter(string nzp_kvar, string num_ls, string num_cnt, string dat_uchet, string dat_pay, decimal val_cnt_old, decimal val_cnt_new, int nzp_serv)
        {
            if (num_cnt == "2.40E+14")
                num_cnt = "2.40001E+14";
            //StreamWriter success = new StreamWriter(@"c:\temp\" + @"tyuiop.txt", true, Encoding.Default);
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billGeu1;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText;
            if (num_cnt != "")
            {
                cmdText = @"SELECT cur_unl, nzp_wp, ist, nzp_counter, nzp_cnttype FROM bill01_data.counters where nzp_serv = " + nzp_serv + " AND nzp_kvar = " + nzp_kvar + " " +
                             "AND num_cnt = '" + num_cnt + "' and dat_uchet = to_date('" + dat_uchet + "','dd-mm-yyyy') and val_cnt = " + val_cnt_old;
            }
            else
            {
                cmdText = @"SELECT cur_unl, nzp_wp, ist, nzp_counter, nzp_cnttype FROM bill01_data.counters where nzp_serv = " + nzp_serv + " AND nzp_kvar = " + nzp_kvar + " " +
                             "AND dat_uchet = to_date('" + dat_uchet + "','dd-mm-yyyy') and val_cnt = " + val_cnt_old;
            }
            //success.WriteLine(cmdText);
            //success.Close();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    if (num_cnt != "")
                    {
                        cmdText = @"SELECT * FROM bill01_data.counters where is_actual = 1 AND nzp_serv = " + nzp_serv + " AND nzp_kvar = " + nzp_kvar + " " +
                             "AND num_cnt = '" + num_cnt + "' and dat_uchet = to_date('" + dat_pay + "','dd-mm-yyyy') and val_cnt = " + val_cnt_new;
                    }
                    else
                    {
                        cmdText = @"SELECT * FROM bill01_data.counters where is_actual = 1 AND nzp_serv = " + nzp_serv + " AND nzp_kvar = " + nzp_kvar + " " +
                             "and dat_uchet = to_date('" + dat_pay + "','dd-mm-yyyy') and val_cnt = " + val_cnt_new;
                    }
                    
                    conn = new NpgsqlConnection(connStr);
                    cmd = new NpgsqlCommand(cmdText, conn);
                    da = new NpgsqlDataAdapter(cmd);
                    DataTable dt2 = new DataTable();
                    da.Fill(dt2);
                    if (dt2.Rows.Count != 0)
                    {
                        return "Показания уже были добавлены";
                    }
                    else
                    {
                        cmdText = @"INSERT INTO bill01_data.counters(nzp_kvar, num_ls, nzp_serv, nzp_cnttype, num_cnt, dat_uchet, val_cnt, is_actual, nzp_user, dat_when, cur_unl, nzp_wp, ist, nzp_counter) 
                              VALUES(" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + dt.Rows[0][4].ToString() + ", '" + num_cnt + "', to_date('" + dat_pay + "','dd-mm-yyyy'), " + val_cnt_new + ", 1, 1, current_date, " + dt.Rows[0][0].ToString()
                                       + ", " + dt.Rows[0][1].ToString() + ", " + dt.Rows[0][2].ToString() + ", " + dt.Rows[0][3].ToString() + ")";
                        cmd = new NpgsqlCommand(cmdText, conn);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                        return "Success";
                    }                 
                }
                else if (dt.Rows.Count > 2)
                {
                    conn.Open();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (num_cnt != "")
                        {
                            cmdText = @"SELECT * FROM bill01_data.counters where is_actual = 1 AND nzp_serv = " + nzp_serv + " AND nzp_kvar = " + nzp_kvar + " " +
                                 "AND num_cnt = '" + num_cnt + "' and dat_uchet = to_date('" + dat_pay + "','dd-mm-yyyy') and val_cnt = " + val_cnt_new;
                        }
                        else
                        {
                            cmdText = @"SELECT * FROM bill01_data.counters where is_actual = 1 AND nzp_serv = " + nzp_serv + " AND nzp_kvar = " + nzp_kvar + " " +
                                 "and dat_uchet = to_date('" + dat_pay + "','dd-mm-yyyy') and val_cnt = " + val_cnt_new;
                        }
                        conn = new NpgsqlConnection(connStr);
                        cmd = new NpgsqlCommand(cmdText, conn);
                        da = new NpgsqlDataAdapter(cmd);
                        DataTable dt2 = new DataTable();
                        da.Fill(dt2);
                        if (dt2.Rows.Count != 0)
                        {
                            return "Показания уже были добавлены";
                        }
                        else
                        {
                            cmdText = @"INSERT INTO bill01_data.counters(nzp_kvar, num_ls, nzp_serv, nzp_cnttype, num_cnt, dat_uchet, val_cnt, is_actual, nzp_user, dat_when, cur_unl, nzp_wp, ist, nzp_counter) 
                              VALUES(" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + dt.Rows[i][4].ToString() + ", '" + num_cnt + "', to_date('" + dat_pay + "','dd-mm-yyyy'), " + val_cnt_new + ", 1, 1, current_date, " + dt.Rows[i][0].ToString()
                                       + ", " + dt.Rows[i][1].ToString() + ", " + dt.Rows[i][2].ToString() + ", " + dt.Rows[i][3].ToString() + ")";
                            cmd = new NpgsqlCommand(cmdText, conn);
                            cmd.ExecuteNonQuery();
                        }                      
                    }
                    conn.Close();
                    return "Success";
                }
                else
                {
                    return "Не найдено не одного счетчика по входным параметрам";
                }
            }
            catch (Exception e)
            {
                return "Ошибка";
            }
        }
    }
}