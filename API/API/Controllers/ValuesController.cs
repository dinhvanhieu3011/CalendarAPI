using API.Models;
using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Web;
using System.Web.Http;

namespace API.Controllers
{
    //[Authorize]
    public class ValuesController : ApiController
    {
        webapiEntities con = new webapiEntities();

        //http://27.71.231.75:82/api/Values/loginStatus?userName=CT010215&passWord=30/11/1995
        [Route("api/Values/loginStatus")]
        [HttpGet]
        public IHttpActionResult loginStatus([FromUri] LoginInfo loginInfo)
        {
    
            ChromeDriver chromeDriver = new ChromeDriver(HttpContext.Current.Server.MapPath("/Chrome/"));
            WebDriverWait wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(10));
            string errorCode = Login(chromeDriver, wait,loginInfo.userName, loginInfo.passWord).ToString();
            chromeDriver.Close();
            chromeDriver.Quit();
        
            if (errorCode == "-1")
            {
                return Ok(new { data = "", errorCode = errorCode });
            }
            else
            {
                var data = con.sinhViens.Where(x => x.userName == loginInfo.userName).FirstOrDefault();
                return Ok(new { data = new
                {
                    MSSV= data.mssv,
                    hoTen=data.ten,
                    khoa= data.khoa,
                    lop = data.lop,
                    trangThaiHocTap = data.hoDem,
                }  , errorCode = errorCode });
            }
        }

        #region Api lấy dữ liệu từ db
        //Lấy điểm trung bình http://27.71.231.75:82/api/Values/getScoreMedium?mssv=CT010215
        [Route("api/Values/getScoreMedium")]
        [HttpGet]
        public IHttpActionResult getScoreMedium([FromUri] string mssv)
        {
            var data = con.diemTrungBinhHocTaps.Where(x => x.mssv == mssv).ToList();
            if (data.Count < 1)
            {
                return Ok(new { data = data, errorCode = 0 });
            }
            else
            {
                return Ok(new { data = data, errorCode = 1 });
            }
        }
        //Lấy điểm chi tiết http://27.71.231.75:82/api/Values/getBangDiemChiTiet?mssv=CT010215
        [Route("api/Values/getBangDiemChiTiet")]
        [HttpGet]
        public IHttpActionResult getBangDiemChiTiet([FromUri] string mssv)
        {
            var data = con.bangDiemChiTiets.Where(x => x.mssv == mssv).ToList();
            if (data.Count < 1)
            {
                return Ok(new { data = data, errorCode = 0 });
            }
            else
            {
                return Ok(new { data = data, errorCode = 1 });
            }
        }

        //Lấy điểm xử lí hoạc vụ http://27.71.231.75:82/api/Values/getXuLyHocVu?mssv=CT010215
        [Route("api/Values/getXuLyHocVu")]
        [HttpGet]
        public IHttpActionResult getXuLyHocVu([FromUri] string mssv)
        {
            var data = con.xuLyHocVus.Where(x => x.mssv == mssv).ToList();
            if (data.Count < 1)
            {
                return Ok(new { data = data, errorCode = 0 });
            }
            else
            {
                return Ok(new { data = data, errorCode = 1 });
            }
        }

        //Lấy điểm lịch học http://27.71.231.75:82/api/Values/getLichHoc?mssv=CT010215
        [Route("api/Values/getLichHoc")]
        [HttpGet]
        public IHttpActionResult getLichHoc([FromUri] string mssv)
        {
            var data = con.lichHocs.Where(x => x.mssv == mssv).ToList();
            if (data.Count < 1)
            {
                return Ok(new { data = data, errorCode = 0 });
            }
            else
            {
                return Ok(new { data = data, errorCode = 1 });
            }
        }

        //Lấy điểm lệ phí học phí http://27.71.231.75:82/api/Values/getlephiHocphi?mssv=CT010215
        [Route("api/Values/getlephiHocphi")]
        [HttpGet]
        public IHttpActionResult getlephiHocphi([FromUri] string mssv)
        {
            var data = con.lePhihocPhis.Where(x => x.mssv == mssv).ToList();
            if (data.Count < 1)
            {
                return Ok(new { data = data, errorCode = 0 });
            }
            else
            {
                return Ok(new { data = data, errorCode = 1 });
            }
        }

        //Lấy thông báo http://27.71.231.75:82/api/Values/getThongBao?mssv=CT010215
        [Route("api/Values/getThongBao")]
        [HttpGet]
        public IHttpActionResult getThongBao([FromUri] string mssv)
        {
            var data = con.thongBaos.Where(x => x.mssv == mssv || x.mssv == "All").ToList();
            if (data.Count < 1)
            {
                return Ok(new { data = data, errorCode = 0 });
            }
            else
            {
                return Ok(new { data = data, errorCode = 1 });
            }
        }
        #endregion

        #region Common Help
        public List<DateTime> GetWeekdayInRange( DateTime from, DateTime to, DayOfWeek day)
        {
            const int daysInWeek = 7;
            var result = new List<DateTime>();
            var daysToAdd = ((int)day - (int)from.DayOfWeek + daysInWeek) % daysInWeek;
            while (from < to.AddDays(-7))
            {
                from = from.AddDays(daysToAdd);
                result.Add(from);
                daysToAdd = daysInWeek;
            } ;
            return result;
        }
        public int CheckExistAccount(string userName, string passWord)
        {
            List<sinhVien> lstSv = con.sinhViens.Where(x => x.userName == userName).ToList();
            if (lstSv.Count == 0)
            {
                // Chưa tồn tại trong db
                return -1;
            }
            else
            {
                if (lstSv[0].passWord != passWord)
                {
                    // Tồn tại trong db
                    return 0;
                }
                else
                    return 1;
            }
        }
        public int Login(ChromeDriver chromeDriver, WebDriverWait wait,string userName, string passWord)
        {
            chromeDriver.Url = "http://qldt.actvn.edu.vn/CMCSoft.IU.Web.info/Login.aspx";
            chromeDriver.Navigate();
            string source = chromeDriver.PageSource;
            // tìm username input
            var userNameInput = chromeDriver.FindElementById("txtUserName");
            userNameInput.SendKeys(userName);
            // tìm password input
            var passWordInput = chromeDriver.FindElementById("txtPassword");
            passWordInput.SendKeys(passWord);
            // tìm nust login
            var btnSubmit = chromeDriver.FindElementById("btnSubmit");
            btnSubmit.Click();
            chromeDriver.FindElement(By.Id("PageHeader1_lblUserRole"));
            wait.Until(driver => driver.FindElement(By.Id("PageHeader1_lblUserRole")).Displayed);
            string currentURL = chromeDriver.Url;
            if (currentURL.ToLower() != "http://qldt.actvn.edu.vn/CMCSoft.IU.Web.info/home.aspx".ToLower())
            {
                //Không đăng nhập thành công(tài khoản mật khẩu sai)
                return -1;
            }
            else
            {
                if (CheckExistAccount(userName, passWord) == -1)
                {
                    #region Vào cái trang thông tin cá nhân này lấy ít thông tin
                    //http://qldt.actvn.edu.vn/CMCSoft.IU.Web.Info/StudentViewExamList.aspx
                    chromeDriver.Url = "http://qldt.actvn.edu.vn/CMCSoft.IU.Web.info/StudentMark.aspx";
                    string hoTen = chromeDriver.FindElementById("lblStudentName").Text;
                    string khoa = chromeDriver.FindElementById("lblAy").Text;
                    string lop = chromeDriver.FindElementById("lblAdminClass").Text;
                    #endregion
                    #region Save Student
                    sinhVien sV = new sinhVien();
                    sV.userName = userName;
                    sV.passWord = passWord;
                    sV.lop = lop;
                    sV.khoa = lop;
                    sV.ten = hoTen;
                    con.sinhViens.Add(sV);
                    con.SaveChanges();
                    #endregion
                    return 2;
                }
                else if (CheckExistAccount(userName, passWord) == 0)
                {
                    return 0;
                }
                else
                {
                    sinhVien sV = con.sinhViens.Where(x => x.userName == userName).FirstOrDefault();
                    sV.passWord = passWord;
                    con.SaveChanges();
                    return 1;
                }
            }
        }
        // Lấy điểm trung bình
        public List<diemTrungBinhHocTap> dongBoDiemTrungBinh(WebDriverWait wait, ChromeDriver chromeDriver, string MSSV)
        {
            chromeDriver.Url = "http://qldt.actvn.edu.vn/CMCSoft.IU.Web.info/StudentMark.aspx";
            wait.Until(driver => driver.FindElement(By.Id("lblStudentName")).Displayed);
            List<diemTrungBinhHocTap> records = new List<diemTrungBinhHocTap>();
            List<IWebElement> rows = chromeDriver.FindElements(By.CssSelector("#grdResult > tbody > tr:nth-child(n)")).ToList();
            foreach (IWebElement row in rows)
            {
                List<IWebElement> cells = row.FindElements(By.CssSelector("td")).ToList();
                diemTrungBinhHocTap record = new diemTrungBinhHocTap();
                record.mssv = MSSV;
                record.namHoc = cells[0].Text.ToString();
                record.hocKy = cells[1].Text.ToString();
                record.value_1 = cells[2].Text.ToString();
                record.value_2 = cells[3].Text.ToString();
                record.value_3 = cells[4].Text.ToString();
                record.value_4 = cells[5].Text.ToString();
                record.value_5 = cells[6].Text.ToString();
                record.value_6 = cells[7].Text.ToString();
                record.value_7 = cells[8].Text.ToString();
                record.value_8 = cells[9].Text.ToString();
                record.value_9 = cells[10].Text.ToString();
                record.value_10 = cells[11].Text.ToString();
                record.value_11 = cells[12].Text.ToString();
                record.value_12 = cells[13].Text.ToString();
                records.Add(record);
            }
            return records;
        }

        public List<bangDiemChiTiet> dongBoBangDiemChiTiet(WebDriverWait wait, ChromeDriver chromeDriver, string MSSV)
        {
            chromeDriver.Url = "http://qldt.actvn.edu.vn/CMCSoft.IU.Web.info/StudentMark.aspx";
            wait.Until(driver => driver.FindElement(By.Id("tblStudentMark")).Displayed);
            List<bangDiemChiTiet> records = new List<bangDiemChiTiet>();
            List<IWebElement> rows = chromeDriver.FindElements(By.CssSelector("#tblStudentMark > tbody > tr:nth-child(n)")).ToList();

            foreach (IWebElement row in rows)
            {
                List<IWebElement> cells = row.FindElements(By.CssSelector("td")).ToList();
                if(cells.Count > 13)
                {
                    bangDiemChiTiet record = new bangDiemChiTiet();
                    record.STT = cells[0].Text.ToString(); ;
                    record.maHocPhan = cells[1].Text.ToString();
                    record.tenHocPhan = cells[2].Text.ToString();
                    record.soTC = cells[3].Text.ToString();
                    record.lanHoc = cells[4].Text.ToString();
                    record.lanThi = cells[5].Text.ToString();
                    record.diemThu = cells[6].Text.ToString();
                    record.laDiemTongKetMon = cells[7].Text.ToString();
                    record.danhGia = cells[8].Text.ToString();
                    record.mssv = MSSV;
                    record.tp1 = cells[10].Text.ToString();
                    record.tp2 = cells[11].Text.ToString();
                    record.thi = cells[12].Text.ToString();
                    record.tkhp2 = cells[13].Text.ToString();
                    records.Add(record);
                }
            }
            return records;
        }

        public List<lichHoc> dongBoLicHoc(WebDriverWait wait, ChromeDriver chromeDriver, string MSSV)
        {
            if (File.Exists(HttpContext.Current.Server.MapPath("/download/ThoiKhoaBieuSinhVien.xls")))
            {
                System.IO.File.Delete(HttpContext.Current.Server.MapPath("/download/ThoiKhoaBieuSinhVien.xls"));
            }

            List<lichHoc> records = new List<lichHoc>();
            chromeDriver.Url = "http://qldt.actvn.edu.vn/CMCSoft.IU.Web.Info/Reports/Form/StudentTimeTable.aspx";
            wait.Until(driver => driver.FindElement(By.Id("drpType")).Displayed);
            ((IJavaScriptExecutor)chromeDriver).ExecuteScript("document.getElementById('drpType').value = 'B'");
            ((IJavaScriptExecutor)chromeDriver).ExecuteScript("document.getElementById('btnView').click()");

            for (var i = 0; i < 30; i++)
            {
                if (File.Exists(HttpContext.Current.Server.MapPath("/download/ThoiKhoaBieuSinhVien.xls"))) { break; }
                Thread.Sleep(1000);
            }

            chromeDriver.Close();
            chromeDriver.Quit();
            if (File.Exists(HttpContext.Current.Server.MapPath("/download/ThoiKhoaBieuSinhVien.xls"))) //helps to check if the zip file is present
            {
                using (var stream = System.IO.File.Open(HttpContext.Current.Server.MapPath("/download/ThoiKhoaBieuSinhVien.xls"), FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                // reader.GetDouble(0);
                            }
                        } while (reader.NextResult());

                        // 2. Use the AsDataSet extension method
                        DataSet result = reader.AsDataSet();
                        DataTable dt = result.Tables[0];
                        //List<string> lstTinh = db.AsEnumerable().Select(x => x["Column0"].ToString()).Distinct().ToList();
                        if (dt.Rows.Count > 2)
                        {

                            foreach (DataRow row in dt.Rows)
                            {
                                if (dt.Rows.IndexOf(row) > 10)
                                {
                                    if(row[0].ToString()!="")
                                    {
                                    
                                        List<DateTime> lstNgayHoc = new List<DateTime>() ;
                                        string caHoc = row[8].ToString();
                                        string monHoc = row[3].ToString();
                                        string maHocPhan = row[1].ToString();
                                        string lopHocPhan = row[4].ToString();
                                        string giaoVien = row[7].ToString();
                                        DateTime f = DateTime.Parse(row[10].ToString().Split('-')[0]);
                                        DateTime t = DateTime.Parse(row[10].ToString().Split('-')[1]);
                                        switch (row[0].ToString())
                                        {
                                            case "2": lstNgayHoc = GetWeekdayInRange(f, t, DayOfWeek.Monday); break;
                                            case "3": lstNgayHoc = GetWeekdayInRange(f, t, DayOfWeek.Tuesday); break;
                                            case "4": lstNgayHoc = GetWeekdayInRange(f, t, DayOfWeek.Wednesday); break;
                                            case "5": lstNgayHoc = GetWeekdayInRange(f, t, DayOfWeek.Thursday); break;
                                            case "6": lstNgayHoc = GetWeekdayInRange(f, t, DayOfWeek.Friday); break;
                                            case "7": lstNgayHoc = GetWeekdayInRange(f, t, DayOfWeek.Saturday); break;
                                        }
                                       for(int i = 0; i < lstNgayHoc.Count; i ++)
                                        {
                                            lichHoc lh = new lichHoc();
                                            lh.mssv = MSSV;
                                            lh.lopHocPhan = lopHocPhan;
                                            lh.caHoc = caHoc;
                                            lh.maHocPhan = maHocPhan;
                                            lh.monHoc = monHoc;
                                            lh.giaoVien = giaoVien;
                                            lh.ngayHoc = lstNgayHoc[i].ToString("dd/MM/yyyy");
                                            records.Add(lh);
                                        }
                                    }
                                    else
                                    {
                                        return records;
                                    }
                                }
                            }
                            return records;
                        }
                        else
                        {
                            return records;
                        }
                    }
                }
            }
            return records;
        }


        #endregion

        #region Clone từ web trường về db server 
        [Route("api/Values/synchronizationMarkAvg")]
        [HttpGet]
        public IHttpActionResult synchronizationMarkAvg([FromUri] LoginInfo loginInfo)
        {
            ChromeDriver chromeDriver = new ChromeDriver(HttpContext.Current.Server.MapPath("/Chrome/"));
            WebDriverWait wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(10));
            try
            {
                string errorCode = Login(chromeDriver, wait, loginInfo.userName, loginInfo.passWord).ToString();
                if (errorCode == "-1")
                {
                    chromeDriver.Close();
                    chromeDriver.Quit();
                    return Ok(new { errorCode = "0" });
                }
                else
                {
                    con.diemTrungBinhHocTaps.RemoveRange(con.diemTrungBinhHocTaps.Where(x => x.mssv == loginInfo.userName));
                    List<diemTrungBinhHocTap> newRecords = dongBoDiemTrungBinh(wait, chromeDriver, loginInfo.userName);
                    chromeDriver.Close();
                    chromeDriver.Quit();
                    con.diemTrungBinhHocTaps.AddRange(newRecords);
                    con.SaveChanges();
                    return Ok(new { errorCode = "1" });
                }
            }
            catch(Exception E)
            {
                return Ok(new { errorCode = "0" });
                chromeDriver.Close();
                chromeDriver.Quit();
            }
        }

        [Route("api/Values/dongBoBangDiemChiTiet")]
        [HttpGet]
        public IHttpActionResult dongBoBangDiemChiTiet([FromUri] LoginInfo loginInfo)
        {

            var chromeDriver = new ChromeDriver(HttpContext.Current.Server.MapPath("/Chrome/"));

            WebDriverWait wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(10));
            try
            {
                string errorCode = Login(chromeDriver, wait, loginInfo.userName, loginInfo.passWord).ToString();
                if (errorCode == "-1")
                {
                    chromeDriver.Close();
                    chromeDriver.Quit();
                    return Ok(new { errorCode = "0" });
                }
                else
                {
                    con.bangDiemChiTiets.RemoveRange(con.bangDiemChiTiets.Where(x => x.mssv == loginInfo.userName));
                    con.SaveChanges();
                    List<bangDiemChiTiet> newRecords = dongBoBangDiemChiTiet(wait, chromeDriver, loginInfo.userName);
                    chromeDriver.Close();
                    chromeDriver.Quit();
                    con.bangDiemChiTiets.AddRange(newRecords);
                    con.SaveChanges();
                    return Ok(new { errorCode = "1" });
                }
            }
            catch (Exception E)
            {
                return Ok(new { errorCode = "0" });
                chromeDriver.Close();
                chromeDriver.Quit();
            }
        }
        [Route("api/Values/dongBoLicHoc")]
        [HttpGet]
        public IHttpActionResult dongBoLicHoc([FromUri] LoginInfo loginInfo)
        {
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", HttpContext.Current.Server.MapPath("/download/"));
            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            var chromeDriver = new ChromeDriver(HttpContext.Current.Server.MapPath("/Chrome/"), chromeOptions);
          
            WebDriverWait wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(10));
            try
            {
                string errorCode = Login(chromeDriver, wait, loginInfo.userName, loginInfo.passWord).ToString();
                if (errorCode == "-1")
                {
                    chromeDriver.Close();
                    chromeDriver.Quit();
                    return Ok(new { errorCode = "0" });
                }
                else
                {
                    con.lichHocs.RemoveRange(con.lichHocs.Where(x => x.mssv == loginInfo.userName));
                    con.SaveChanges();
                    List<lichHoc> newRecords = dongBoLicHoc(wait, chromeDriver, loginInfo.userName);

                    con.lichHocs.AddRange(newRecords);
                    con.SaveChanges();
                    return Ok(new { errorCode = "1" });
                }
            }
            catch (Exception E)
            {
                return Ok(new { errorCode = "0" });
                chromeDriver.Close();
                chromeDriver.Quit();
            }
        }
        #endregion
    }
}
