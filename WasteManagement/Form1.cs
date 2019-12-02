using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WasteManagement
{
    public partial class Form1 : Form
    {
        DataTable WasteTable = new DataTable();
        DataTable DataOfCellTable = new DataTable();//טבלת פירוט על כל תא של קודי פסולת
        Waste WasteObject = new Waste();
        bool Start=true;//התחלה או לא
        int LeftOver = 0;//משתנה פירוט פסולת,אומר כמה נשאר לפרט
        bool CboCodeWaste = false;//האם קומבובוקס פירוט קוד פסולת השתנה,עבור שחרור כפתור אשר
        //משתני עדכון
        bool CellUpdated = false;//בדיקה אם מדובר ברשומה שעודכנה או שלא
        int eRowIndex ,eColumnIndex;//לדעת באיזה תא מדובר בשביל אחרי העדכון לצבוע אותו בטורקיז
        bool SaveData = false;//יחסום צביעת תא אם לא שמרנו נתונים
        bool RealeseComboWasteCode=false;// מונע שגיאה של קומבובוקס קודי פסולת
        Dictionary<int, int> SumScrap = new Dictionary<int, int>();//סכום פסולות עבור יום בודד-לאקסל
        /// <summary>
        /// תיאור מפורט של פסולת בטב עדכון
        /// </summary>
        DataTable DescriptionWasteTable;


        public Form1()
        {
            InitializeComponent();
            
            ShowWaste();
            Start = false;//אחרי שסיימנו להציג, ההצגה הבאה תהיה רק לפי בחירת חודשים לפי בחירת משתמש
            StartPositionFunc();

            DataTable GetWasteCodes = WasteObject.GetWasteCodes();
            cbo_CodeWaste.DataSource = GetWasteCodes;
            cbo_CodeWaste.DisplayMember = "code";
            cbo_CodeWaste.SelectedIndex = -1;
            RealeseComboWasteCode = true;//כעת אפשר לההשתמש בפונקציה של שינוי טקסט בקומבובוקס


        }


        /// <summary>
        /// מסך התחלתי
        /// </summary>
        private void StartPositionFunc()
        {
            for (int i = DateTime.Now.Year; i >= 2010; i--)
            {
                Cbo_YearReport.Items.Add(i);
            }
            Cbo_YearReport.SelectedIndex = 0;
            Cbo_MonthReport.Items.AddRange(CultureInfo.CurrentCulture.DateTimeFormat.MonthNames);//הוספת שמות חודשים
            Cbo_MonthReport.SelectedIndex = DateTime.Now.Month- 1;//חודש נוכחי בקומבובוקס
            lbl_LastUpdate.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
        }

        /// <summary>
        /// הצגת טבלת פסולת לפי חודשים
        /// </summary>
        private void ShowWaste()
        {
            string MonthNumber;
            int Days,Year;

            //בתחילת התוכנית הצגה של החודש הנוכחי
            if (Start)
            {
                 MonthNumber = DateTime.Now.Month.ToString();
                if (int.Parse(MonthNumber) < 10)
                    MonthNumber = "0" + MonthNumber;//הוספת 0 לחודש חד ספרתי
                Days = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);//כמה ימים בחודש העכשווי
                WasteTable = WasteObject.GetWasteList(MonthNumber,DateTime.Now.Year,Days,true);//דיפולט-חודש ושנה של היום
                lbl_MonthReport.Text = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month) + " " + DateTime.Now.Year;
            }

            else
            {
                string Month = Cbo_MonthReport.SelectedItem.ToString();
                //בניית תאריך רצוי עבור השאילתה
                Year = int.Parse(Cbo_YearReport.SelectedItem.ToString());//שנה וחודש לפי בחירת המשתמש
                MonthNumber = DateTime.ParseExact(Month, "MMMM", CultureInfo.CurrentCulture).Month.ToString();
                if (int.Parse(MonthNumber) < 10)
                    MonthNumber = "0" + MonthNumber;//הוספת 0 לחודש חד ספרתי
                Days= DateTime.DaysInMonth(Year,int.Parse(MonthNumber));//כמה ימים בחודש הנבחר
                WasteTable = WasteObject.GetWasteList(MonthNumber, Year,Days,false);
                lbl_MonthReport.Text = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(int.Parse(MonthNumber)) + " " + Year;
            }
     
            DgvWasteMonth.DataSource = WasteTable;

            //מילוי המספרים למטה
            lbl_totalCured.Text = Convert.ToDecimal(WasteObject.SumProduction).ToString("#,#");
            //WasteObject.GetSumWaste();
            lbl_totalScrap.Text= Convert.ToDecimal(WasteObject.SumWaste).ToString("#,#");       
            lbl_EntryToWarehouse.Text = Convert.ToDecimal(WasteObject.EntryToWarehouse).ToString("#,#");//כמה ק"ג מגופרים נכנסו למחסן
            lbl_ScrapTGT.Text = Convert.ToDecimal(WasteObject.ScrapTgtPrecent).ToString()+"%";//מטרת אחוזי פסולת כוללים מתוך ייצור נטו
            lbl_PrecentTotalScrap.Text= Convert.ToDecimal(WasteObject.ScarpTotalPrecent).ToString("p2");//כמה אחוזי פסולת בפועל
            if(WasteObject.ScrapTgtPrecent<WasteObject.ScarpTotalPrecent*100)
                lbl_PrecentTotalScrap.ForeColor = System.Drawing.Color.Red;
            else
                lbl_PrecentTotalScrap.ForeColor = System.Drawing.Color.GreenYellow;

            //צביעת תאים שעודכנו לפי בדיקה מול טבלת MSVQTP בs400
            DataOfCellTable = WasteObject.GetCellsData();

            ////עיצוב טבלה
            DgvWasteMonth.Columns["Total"].DefaultCellStyle.BackColor = Color.Yellow;//צביעה באדום טור טוטל
            this.DgvWasteMonth.Columns["Catalog Number"].Frozen = true;
            this.DgvWasteMonth.Columns["description"].Frozen = true;

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.Fixed3D; ;
            this.WindowState = FormWindowState.Maximized;
        }


        /// <summary>
        /// שינוי תצוגה לפי תאריך
        /// </summary>
        private void btn_show_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            DgvWasteMonth.DataSource = null;
            WasteObject = new Waste();
            ShowWaste();
            Cursor.Current = Cursors.Default;
        }


        /// <summary>
        /// פתיחת חלונית עדכון פסולת
        /// </summary>
        private void DgvWasteMonth_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == DgvWasteMonth.Columns["Total"].DisplayIndex
                || e.ColumnIndex == DgvWasteMonth.Columns["Mtd Scrap% Warehouse"].DisplayIndex || e.ColumnIndex == DgvWasteMonth.Columns["Scrap TGT %"].DisplayIndex) return; //רק בלחיצה על מספר ספציפי יהיה אפשר לעדכן
            eRowIndex = e.RowIndex;//אחרי העדכון ייצבע התא בטורקיז
            eColumnIndex = e.ColumnIndex;

            tabControl1.SelectedTab = WasteUpdate;//מעבר לעדכון 

            WasteObject.CatalogNumber = DgvWasteMonth.Rows[e.RowIndex].Cells["Catalog Number"].Value.ToString();
            WasteObject.DescriptionCatalog = DgvWasteMonth.Rows[e.RowIndex].Cells["description"].Value.ToString();
            string AmountWaste = DgvWasteMonth.Rows[e.RowIndex].Cells["Total"].Value.ToString().Replace(",", "");//כמות טוטלית של המק"ט שבחרנו
            WasteObject.AmountWaste = int.Parse(AmountWaste);//כמות טוטלית של המק"ט שבחרנו עדכון במחלקה
            lbl_WasteUpdate.Text = WasteObject.CatalogNumber + "-" + WasteObject.DescriptionCatalog; ;//מספר קטלוגי ותיאור
            lbl_dateWantedUpdate.Text = DgvWasteMonth.Columns[e.ColumnIndex].HeaderText + "/" + WasteObject.MonthNumber + "/" + WasteObject.Year;//תאריך נבחר לעדכון
            WasteObject.AmountPerDay = int.Parse(DgvWasteMonth.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());//כמה יש מאותו יום שאנחנו רוצים לעדכן בו כמות
            lbl_AmountWaste.Text = Convert.ToDecimal(WasteObject.AmountPerDay).ToString("#,#");//כמות פסולת כוללת מתוך אותו מק"ט באותו יום
            LeftOver = WasteObject.AmountPerDay;//כמה נשאר לפרט,בהתחלה נשאר הכל
            lbl_totalLeftover.Text = Convert.ToDecimal(LeftOver).ToString("#,#");

            //טבלה חדשה עבור פירוט כמויות פסולת
            DescriptionWasteTable = new DataTable();
            DescriptionWasteTable.Columns.Add("סידורי");
            DescriptionWasteTable.Columns.Add("כמות");
            DescriptionWasteTable.Columns.Add("קוד");
            DataGrid_Desc.DataSource = DescriptionWasteTable;
            this.DataGrid_Desc.Columns["סידורי"].ReadOnly = true;
            this.DataGrid_Desc.Columns["קוד"].ReadOnly = true;

            if (DgvWasteMonth.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == Color.Turquoise)//אם מדובר בתא שמעודכן כבר-מעודכן צבוע בטורקיז
            {
                int RowsUpdate = 0;//כמה שורות בטבלה
                for (int i = 0; i < DataOfCellTable.Rows.Count; i++)
                {
                    if (WasteObject.CatalogNumber == DataOfCellTable.Rows[i]["MSMKT"].ToString())//מק"ט רלוונטי
                    {
                        DataRow row = DescriptionWasteTable.NewRow();//הכנסה לטבלה את כל העדכונים שנעשו
                        DescriptionWasteTable.Rows.Add(row);
                        DescriptionWasteTable.Rows[RowsUpdate]["כמות"] = int.Parse(double.Parse(DataOfCellTable.Rows[i]["MSQTY"].ToString()).ToString());
                        DescriptionWasteTable.Rows[RowsUpdate]["קוד"] = DataOfCellTable.Rows[i]["MSKOD"].ToString();
                        DescriptionWasteTable.Rows[RowsUpdate]["סידורי"] = RowsUpdate + 1;
                        LeftOver -= int.Parse(double.Parse(DataOfCellTable.Rows[i]["MSQTY"].ToString()).ToString());//מעדכן כמה נשאר מהטוטל
                        RowsUpdate++;
                    }
                }
                DataGrid_Desc.DataSource = DescriptionWasteTable;
                lbl_totalLeftover.Text = LeftOver.ToString();
                btn_SaveData.Enabled = true;
                RowsUpdate = 0;
                CellUpdated = true;//מעדכן לרשומה שעודכנה במידה וירצה לשנות אותה באירוע cellEditEnd
                this.DataGrid_Desc.Columns["סידורי"].ReadOnly = false;
                this.DataGrid_Desc.Columns["קוד"].ReadOnly = false;
            }

        }

        private void cbo_CodeWaste_SelectedIndexChanged(object sender, EventArgs e)
        {
            CboCodeWaste = true;
            string value;
            if (RealeseComboWasteCode)
            {
                WasteObject.CodeAndDescWaste.TryGetValue(int.Parse(cbo_CodeWaste.Text), out value);
                txt_CodeWaste.Text = value.Trim();
            }

            //int X = cbo_CodeWaste.Location.X+17;
            //int Y= cbo_CodeWaste.Location.Y+160;
            //ComboBox comboBox = new ComboBox();
            //comboBox.Location = new Point(X, Y);
            //comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //comboBox.Font = new System.Drawing.Font("Courier New", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            //comboBox.FormattingEnabled = true;
            //comboBox.Size = new System.Drawing.Size(147, 30);
            //comboBox.RightToLeft=new RightToLeft();
            ////comboBox.Anchor= System.Windows.Forms.AnchorStyles.Top;
            //this.Controls.Add(comboBox);
            //comboBox.BringToFront();

        }

        /// <summary>
        /// כפתור אישור של ציון כמות מסוימת מתוך פסולת
        /// </summary>
        private void btn_confirm_Click(object sender, EventArgs e)
        {
            if (CboCodeWaste == false || txt_AmountSpesific.Text == "") return;
            DataRow row = DescriptionWasteTable.NewRow();
            DescriptionWasteTable.Rows.Add(row);
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["סידורי"] = DescriptionWasteTable.Rows.Count.ToString();
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count-1]["כמות"]= txt_AmountSpesific.Text;
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count-1]["קוד"] = cbo_CodeWaste.Text;
          
            DataGrid_Desc.DataSource = DescriptionWasteTable;

            //חישוב כמה נשאר
            LeftOver -= int.Parse(DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["כמות"].ToString());
            if (LeftOver == 0)
                lbl_totalLeftover.Text = "0";
            else
                lbl_totalLeftover.Text = Convert.ToDecimal(LeftOver).ToString("#,#");
            if(LeftOver<0) lbl_totalLeftover.ForeColor= System.Drawing.Color.Red;

            btn_SaveData.Enabled = true;//כעת ניתן לשמור נתונים
        }


        /// <summary>
        /// כאשר המשתמש רוצה לתקן רשומה בודדת בפירוט הפסולות אחרי שהכניס נתונים כבר
        /// </summary>
        private void Data_Desc_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex == 0 || e.ColumnIndex == 2) return;
            LeftOver = WasteObject.AmountPerDay;//חזרה מאפס כמה יש מאותו יום שאנחנו רוצים לעדכן בו כמות פסולות
            for (int i = 0; i < DataGrid_Desc.RowCount - 1; i++)
            {
                LeftOver -= int.Parse(DataGrid_Desc.Rows[i].Cells["כמות"].Value.ToString());
            }
            lbl_totalLeftover.Text = Convert.ToDecimal(LeftOver).ToString("#,#");
            if (LeftOver < 0) lbl_totalLeftover.ForeColor = System.Drawing.Color.Red;

        }

        private void txt_AmountSpesific_TextChanged(object sender, EventArgs e)
        {
            btn_confirm.Enabled = true;
        }

        /// <summary>
        /// שמירת נתונים לs400 
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (LeftOver < 0)//אם הפירוט יותר מהמכסה הקיימת
                {
                    MessageBox.Show("כמות פירוט פסולת לא תואמת את הכמות הכוללת");
                    return;
                }

                //אם מעדכנים רשומה שהיא הייתה מעודכנת כבר
                if (CellUpdated)
                    WasteObject.DeleteCellUpdate();

                DialogResult d = MessageBox.Show("האם אתה בטוח שברצונך לשמור את הנתונים?", "אישור נתונים", MessageBoxButtons.YesNo);
                if (d == DialogResult.Yes)
                {
                    UpdateCell();//עדכון תא בs400
                    SaveData = true;//לצביעת תא לטורקיז
                    MessageBox.Show("נתונים נשמרו בהצלחה");
                    tabControl1.SelectedTab = Report;
                    CellUpdated = true;
                    //צביעת תאים שעודכנו לפי בדיקה מול טבלת MSVQTP בs400
                    DataOfCellTable = WasteObject.GetCellsData();
                }

                if (d == DialogResult.No)
                    return;
               
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
                //שמירת נתונים

        }

        /// <summary>
        /// עדכון של תא
        /// </summary>
        public void UpdateCell()
        {
            DBService dbs = new DBService();
            string InsertValues = "";
            //שמירה לs400
            for (int i = 0; i < DataGrid_Desc.RowCount - 1; i++)
            {
                //לפי הסדר: מספר קטלוגי,תאריך דיווח,קוד תקלה,כמות,שם מדווח(כרגע ריק) ,תאריך עדכון היום
                InsertValues = $@"insert into  MSK.MSVQTP values ('{WasteObject.CatalogNumber.Trim()}',{DateTime.Parse(lbl_dateWantedUpdate.Text).ToString("yyyyMMdd")}
                                         ,{DataGrid_Desc.Rows[i].Cells["קוד"].Value.ToString()},{DataGrid_Desc.Rows[i].Cells["כמות"].Value.ToString()},'',{DateTime.Now.ToString("yyyyMMdd")})";// ('test11',123456,123,100,'test1',87654321)"
                dbs.executeInsertQuery(InsertValues);
            }
        }

        private void MakeExcellRepBTN_Click(object sender, EventArgs e)
        {
            SendWasteReport();
            //Cursor.Current = Cursors.WaitCursor;
            //Excel.Application xlexcel;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;

            //object misValue = System.Reflection.Missing.Value;
            //xlexcel = new Excel.Application();        
            //xlWorkBook = xlexcel.Workbooks.Add(misValue);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //Excel.Worksheet active = (Excel.Worksheet)xlexcel.ActiveSheet;
            //active.DisplayRightToLeft = false;

            ////יש 40 טורים
            //Excel.Range ColumnB = xlWorkSheet.get_Range("B1");
            //ColumnB.ColumnWidth = 20;
            //xlWorkSheet.get_Range("c1").ColumnWidth = 27;
            //xlWorkSheet.get_Range("A:A").EntireColumn.Hidden = true;
            //xlWorkSheet.get_Range("A5").EntireRow.Hidden = true;

            ////כותרת
            //xlWorkSheet.Cells[3, 2] = "Report Scrap";
            ////xlWorkSheet.get_Range("B3:B3").Merge();//מיזוג תאים
            //xlWorkSheet.get_Range("B3:B3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.get_Range("B3:B3").Font.Bold = true;
            //xlWorkSheet.get_Range("B3:B3").Font.Size = 16;
            //xlWorkSheet.get_Range("B3:B3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("B3:B3").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות

            //xlWorkSheet.Cells[7, 2] = "Topic subject";
            //xlWorkSheet.Cells[7, 2].Font.Size = 14;
            //xlWorkSheet.Cells[7, 2].Font.Bold = true;
            //xlWorkSheet.get_Range("B7:AL7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Cells[7, 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            //xlWorkSheet.get_Range("B6:AL6").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);//צביעה בורוד עמודות טורים
            ////CultureInfo ci = new CultureInfo("en-US");

            ////טוטל נכון לחודש
            //string Month=CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(int.Parse(WasteObject.MonthNumber)).ToString( CultureInfo.InvariantCulture);

            //xlWorkSheet.get_Range("AI4:AL5").Merge();//מיזוג תאים
            //xlWorkSheet.get_Range("AI4:AL5").Font.Bold = true;
            //xlWorkSheet.get_Range("AI4:AL5").Font.Size = 16;
            //xlWorkSheet.get_Range("AI4:AL5").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
            //xlWorkSheet.get_Range("AI4:AL5").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט

            ////תאריך עדכון אחרון
            //xlWorkSheet.Cells[4, 2] = "Month Report";
            //xlWorkSheet.Cells[4, 2].Font.Size = 16;
            //xlWorkSheet.get_Range("A4").RowHeight = 50;
            //xlWorkSheet.Cells[4, 3] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            //if (WasteObject.ThisMonth)
            //    xlWorkSheet.Cells[4, "C"] = $@"נכון ל{DateTime.Today.AddDays(-1).ToShortDateString()}";
            //else
            //    xlWorkSheet.Cells[4, "C"] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            //xlWorkSheet.Cells[4, 3].Font.Size = 16;
            //xlWorkSheet.get_Range("C4:C4").Merge();//מיזוג תאים
            //xlWorkSheet.get_Range("B4:C4").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.get_Range("B4:C4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("B4:C4").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            //xlWorkSheet.get_Range("B4:C4").Font.Bold = true;




            ////הכנסת כותרות של פריטים ל2 עמודות ראשונות
            //for (int i = 0; i < 2; i++)//מתחיל מאחרי הכותרות
            //{
            //    for (int j = 0; j < DgvWasteMonth.Rows.Count; j++)//
            //    {
            //        xlWorkSheet.Cells[9 + j, i+2] = DgvWasteMonth.Rows[j].Cells[i].Value.ToString();
            //    }
            //}

            ////העתקת גריד ויו PER DAYS
            //Dictionary<int, double> DictionaryWHentrence = new Dictionary<int, double>();
            //Dictionary<int, double> DictionarySumCuring = new Dictionary<int, double>();
            //DictionarySumCuring = WasteObject.BringSumCuring();
            //DictionaryWHentrence = WasteObject.BringWarehouseEntrance();//סכומי כניסה למחסן

            //int RowsTotalPerDay = 0;//מספר שורה באקסל של הסיכומים

            /////הסיכומים לרמת הימים
            ////מתחיל אחרי הכותרות
            //for (int i = 2; i < DgvWasteMonth.Columns.Count - 4; i++)
            //{
            //    int SumScrap = 0, j ;
            //    for (j = 0; j < DgvWasteMonth.Rows.Count; j++)//לא כולל טוטל ויעדים
            //    {
            //        xlWorkSheet.Cells[9 + j, i + 6] = DgvWasteMonth.Rows[j].Cells[i].Value.ToString();//I+6=COLUMN G ,J+9=ROW 9 START
            //        bool isNumeric = int.TryParse(DgvWasteMonth.Rows[j].Cells[i].Value.ToString(), out int n);
            //        if(isNumeric)
            //        {
            //            SumScrap += n;
            //        }
            //    }
            //    if (i == 2)//כותרות סיכומים
            //    {
            //        RowsTotalPerDay = 9 + j;
            //        xlWorkSheet.Cells[RowsTotalPerDay, "B"] = "Total Scrap";
            //        xlWorkSheet.Cells[RowsTotalPerDay + 1, "B"] = " Warehous Entrances";
            //        xlWorkSheet.Cells[RowsTotalPerDay + 2, "B"] = "Total Cured";
            //    }

            //    ////סכום פסולות יומי
            //    if (SumScrap != 0)
            //        xlWorkSheet.Cells[RowsTotalPerDay, i + 6] = SumScrap;
            //    ////סכום כניסה למחסן יומי
            //    if (DictionaryWHentrence.ContainsKey(int.Parse(DgvWasteMonth.Columns[i].HeaderText)))//אם היה בכלל הכנסה למחסן באותו יום ספציפי
            //    {
            //        DictionaryWHentrence.TryGetValue(int.Parse(DgvWasteMonth.Columns[i].HeaderText), out double value);
            //        xlWorkSheet.Cells[RowsTotalPerDay + 1, i + 6] = value;
            //    }
            //    //if (i<TableWarehouse.Rows.Count)
            //    //xlWorkSheet.Cells[RowsTotalPerDay+1, i + 4] = TableWarehouse.Rows[i]["Sum"].ToString();//j+1 שורה מתחת,i+4 start column D
            //    ////סכום גיפור יומי
            //    if (DictionarySumCuring.ContainsKey(int.Parse(DgvWasteMonth.Columns[i].HeaderText)))
            //    {
            //        DictionarySumCuring.TryGetValue(int.Parse(DgvWasteMonth.Columns[i].HeaderText), out double value);
            //        xlWorkSheet.Cells[RowsTotalPerDay + 2, i + 6] = value;
            //    }
            //}

            
            ////עיצוב דטה גריד
            //int RowsBorder = DgvWasteMonth.RowCount+11;//-יתחיל בשורה 9 גבולות תא של דטה גריד באקסל
            //xlWorkSheet.get_Range("B6"+":AL"+RowsBorder).Borders.Color = System.Drawing.Color.Black.ToArgb();
            //xlWorkSheet.get_Range("B6" + ":AL" + RowsBorder).Font.Bold = true;
            //xlWorkSheet.get_Range("D6" + ":AL" + RowsBorder).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("B6" + ":AL" + RowsBorder).Font.Size = 12;
            //xlWorkSheet.get_Range("B6" + ":AL" + RowsBorder).RowHeight = 26;
            //xlWorkSheet.get_Range("A6:AK7").RowHeight = 35;//גובה שורה של העמודות
          
            ////צביעת שלושת השורות של הסיכומים
            //xlWorkSheet.get_Range("B"+RowsTotalPerDay+":AL"+(RowsTotalPerDay+2)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.get_Range("B" + RowsTotalPerDay + ":AL" + (RowsTotalPerDay + 2)).Font.Bold = true;
            //xlWorkSheet.get_Range("B" + RowsTotalPerDay + ":AL" + (RowsTotalPerDay + 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט


            ////הקפאת עמודות
            //xlWorkSheet.get_Range("A:D").Application.ActiveWindow.FreezePanes = true;


            ////יצירת כותרות עמודות
            //int Column = 8;//יתחיל מטור 3
            //for (int i = 1; i <= WasteObject.Days; i++)//מספר הימים בחודש יגדיר עמודות יתחיל מטור 3 
            //{
            //    xlWorkSheet.Cells[6, Column] = i;//יום 1 יכניס בשורה 6 טור 3
            //    xlWorkSheet.Cells[7, Column] = "Weight";
            //    Column++;
            //}


            //// טורי טוטל יעדים ועמידה ביעדים
            //xlWorkSheet.Cells[6, "D"] = "TOTAL";
            //xlWorkSheet.get_Range("D6:D7").Merge();
            //xlWorkSheet.get_Range("D6:D7").Font.Size = 11;
            //xlWorkSheet.get_Range("D6:D7").Font.Bold = true;
            //xlWorkSheet.get_Range("D6:D7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("D6:D7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            //xlWorkSheet.get_Range("D6:E7").WrapText = true;
            //xlWorkSheet.get_Range("D6:D7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Cells[6, "E"] = "Scrap % On cured Gross Tonnage";
            //xlWorkSheet.get_Range("E6:E7").Merge();
            //xlWorkSheet.get_Range("E6:E7").Font.Size = 11;
            //xlWorkSheet.get_Range("E6:E7").Font.Bold = true;
            //xlWorkSheet.get_Range("E6:E7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("E6:E7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            //xlWorkSheet.get_Range("E6:E7").WrapText = true;
            //xlWorkSheet.get_Range("E6:E7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Cells[6, "F"] = "Mtd Scrap% Warehouse Tonnage";
            //xlWorkSheet.get_Range("F6:F7").Merge();
            //xlWorkSheet.get_Range("F6:F7").Font.Size=11;
            //xlWorkSheet.get_Range("F6:F7").Font.Bold=true;
            //xlWorkSheet.get_Range("F6:F7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("F6:F7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            //xlWorkSheet.get_Range("F6:F7").WrapText = true;
            //xlWorkSheet.get_Range("F6:F7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Cells[6, "G"] = "Scrap TGT %";
            //xlWorkSheet.get_Range("G6:G7").Merge();
            //xlWorkSheet.get_Range("G6:G7").Font.Size = 11;
            //xlWorkSheet.get_Range("G6:G7").Font.Bold = true;
            //xlWorkSheet.get_Range("G6:G7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            //xlWorkSheet.get_Range("G6:G7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            //xlWorkSheet.get_Range("G6:G7").WrapText = true;
            //xlWorkSheet.get_Range("G6:G7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.get_Range("D6:G" + RowsBorder).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות בולטים להבלטת טוטל ויעדים
            //xlWorkSheet.get_Range("D6:G" + RowsBorder).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;

            ////צביעת תאי טורי יעדים אדום ירוק וצהוב
            //for (int i = 0; i < DgvWasteMonth.RowCount; i++)
            //{
            //    if (!string.IsNullOrEmpty(DgvWasteMonth.Rows[i].Cells["Scrap TGT %"].Value.ToString() as string))
            //        xlWorkSheet.Cells[i + 9, "G"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);//צביעה בצהוב מטרות יעדים דטה גריד ויו מתחיל בשורה 9
            //    //צביעת תאים באדום או ירוק לפי בדיקה של הצבעים בדטה גריד
            //    if (DgvWasteMonth.Rows[i].Cells["Mtd Scrap% Warehouse  TONNAGE"].Style.BackColor==Color.YellowGreen)
            //        xlWorkSheet.Cells[i + 9, "F"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            //    if (DgvWasteMonth.Rows[i].Cells["Mtd Scrap% Warehouse  TONNAGE"].Style.BackColor == Color.Red)
            //        xlWorkSheet.Cells[i + 9, "F"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //}

            //double SumPrecentCured = 0, SumPrecentWH = 0, SumPrecentTarget = 0;
            //int r = 0;
            ////טוטל כולל עמודה D
            ////הכנסת סיכומים (הטורים הכללים)לעמודות הנכונות
            //for (int i = 0; i < DgvWasteMonth.Rows.Count; i++)
            //{
            //    int counter = 0;
            //    for (int j = DgvWasteMonth.Columns.Count; j > DgvWasteMonth.Columns.Count - 4; j--)
            //    {
            //        xlWorkSheet.Cells[i + 9, 7 - counter] = DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString();
            //        counter++;
            //        if (!string.IsNullOrEmpty(DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString()))
            //        {
            //            switch (DgvWasteMonth.Columns[j - 1].HeaderText)//סכימה של טורי הסיכומים
            //            {
            //                case "Scrap % On cured Gross Tonnage":
            //                    SumPrecentCured += double.Parse(DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString().Substring(0, DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString().IndexOf("%")));
            //                    break;

            //                case "Mtd Scrap% Warehouse  TONNAGE":
            //                    SumPrecentWH += double.Parse(DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString().Substring(0, DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString().IndexOf("%")));
            //                    break;

            //                case "Scrap TGT %":
            //                        SumPrecentTarget += double.Parse(DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString().Substring(0, DgvWasteMonth.Rows[i].Cells[j - 1].Value.ToString().IndexOf("%")));
            //                    break;
            //            }
            //        }
            //    }
            //    r = i+1;
            //}
            //xlWorkSheet.Cells[9 + r, "E"] = (SumPrecentCured/100).ToString("P2");
            //xlWorkSheet.Cells[9 + r, "F"] = (SumPrecentWH/100).ToString("P2");
            //xlWorkSheet.Cells[9 + r, "G"] = (SumPrecentTarget/100).ToString("P2");

            //int RowEndTable = RowsBorder+1;
            //RowsBorder++;
            ////הכנסת סיכומים
            //xlWorkSheet.Cells[RowsBorder, "E"] = lblTotalCured.Text;
            //xlWorkSheet.Cells[RowsBorder, "F"] = lbl_totalCured.Text;
            //xlWorkSheet.get_Range("D"+RowsBorder+":E"+RowsBorder).Merge();
            //RowsBorder++;
            //xlWorkSheet.Cells[RowsBorder, "E"] = lblheaderTotalScrap.Text;
            //xlWorkSheet.Cells[RowsBorder, "F"] = lbl_totalScrap.Text;
            //xlWorkSheet.get_Range("D" + RowsBorder + ":E" + RowsBorder).Merge();
            //RowsBorder++;
            //xlWorkSheet.Cells[RowsBorder, "E"] = lblheaderEntryToWarehouse.Text;
            //xlWorkSheet.Cells[RowsBorder, "F"] = lbl_EntryToWarehouse.Text;
            //xlWorkSheet.get_Range("D" + RowsBorder + ":E" + RowsBorder).Merge();
            //RowsBorder++;
            //xlWorkSheet.Cells[RowsBorder, "E"] = lblheaderScrapTgt.Text;
            //xlWorkSheet.Cells[RowsBorder, "F"] = lbl_ScrapTGT.Text;
            //xlWorkSheet.get_Range("D" + RowsBorder + ":E" + RowsBorder).Merge();
            //RowsBorder++;
            //xlWorkSheet.Cells[RowsBorder, "E"] = lblheaderScrapPrecent.Text;
            //xlWorkSheet.Cells[RowsBorder, "F"] = lbl_PrecentTotalScrap.Text;
            //xlWorkSheet.get_Range("D" + RowsBorder + ":E" + RowsBorder).Merge();           
            //if (lbl_PrecentTotalScrap.ForeColor == Color.Red)
            //    xlWorkSheet.Cells[RowsBorder, "F"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //else
            //    xlWorkSheet.Cells[RowsBorder, "F"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);

            
            //Excel.Range chartTotalRange= xlWorkSheet.get_Range("D"+RowEndTable, "F"+RowsBorder);
            //chartTotalRange.Font.Size = 14;
            //chartTotalRange.Font.Bold = true;
            //chartTotalRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot;//גבולות בולטים להבלטת טוטל ויעדים
            //chartTotalRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot;//גבולות בולטים להבלטת טוטל ויעדים
            //chartTotalRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot;
            ////xlWorkSheet.get_Range("E1:F1").ColumnWidth = 35;



            //xlexcel.Visible = true;
            //Cursor.Current = Cursors.Default;
            ////xlWorkBook.Close();
            //releaseObject(xlexcel);
            //releaseObject(xlWorkBook);
            //releaseObject(xlWorkSheet);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }



        private void SendWasteReport()
        {
            Cursor.Current = Cursors.WaitCursor;
            Excel.Application xlexcel;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlWorkBook = xlexcel.Workbooks.Add(misValue);         
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Worksheet active = (Excel.Worksheet)xlexcel.ActiveSheet;
            active.DisplayRightToLeft = false;

            //יש 40 טורים
            Excel.Range ColumnB = xlWorkSheet.get_Range("B1");
            ColumnB.ColumnWidth = 20;
            xlWorkSheet.get_Range("c1").ColumnWidth = 32;
            xlWorkSheet.get_Range("g1").ColumnWidth = 11.75;
            xlWorkSheet.get_Range("A:A").EntireColumn.Hidden = true;
            xlWorkSheet.get_Range("A5").EntireRow.Hidden = true;
            xlWorkSheet.get_Range("a1").EntireRow.Hidden = true;
            xlWorkSheet.get_Range("a2").EntireRow.Hidden = true;
            xlWorkSheet.get_Range("A8").EntireRow.Hidden = true;
            xlWorkSheet.get_Range("B8").EntireRow.Hidden = true;


            //B3
            //כותרת
            xlWorkSheet.Cells[3, 2] = "Report Scrap";
            //xlWorkSheet.get_Range("B3:B3").Merge();//מיזוג תאים
            xlWorkSheet.get_Range("B3:B3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.get_Range("B3:B3").Font.Bold = true;
            xlWorkSheet.get_Range("B3:B3").Font.Size = 16;
            xlWorkSheet.get_Range("B3:B3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("B3:B3").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            //B7
            xlWorkSheet.Cells[7, 2] = "Topic subject";
            xlWorkSheet.Cells[7, 2].Font.Size = 14;
            xlWorkSheet.Cells[7, 2].Font.Bold = true;
            xlWorkSheet.get_Range("B7:AL7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.Cells[7, 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.get_Range("B6:AL6").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);//צביעה בורוד עמודות טורים
            //CultureInfo ci = new CultureInfo("en-US");

            //טוטל נכון לחודש
            string Month = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(int.Parse(WasteObject.MonthNumber)).ToString(CultureInfo.InvariantCulture);

            //B4-C4
            //תאריך עדכון אחרון
            xlWorkSheet.Cells[4, 2] = "Month Report";
            xlWorkSheet.Cells[4, 2].Font.Size = 16;
            xlWorkSheet.get_Range("A4").RowHeight = 50;
            xlWorkSheet.Cells[4, 3] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            if (WasteObject.ThisMonth)
                xlWorkSheet.Cells[4, "C"] = $@"נכון ל{DateTime.Today.AddDays(-1).ToShortDateString()}";
            else
                xlWorkSheet.Cells[4, "C"] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            xlWorkSheet.Cells[4, 3].Font.Size = 16;
            xlWorkSheet.get_Range("C4:C4").Merge();//מיזוג תאים
            xlWorkSheet.get_Range("B4:C4").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.get_Range("B4:C4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("B4:C4").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            xlWorkSheet.get_Range("B4:C4").Font.Bold = true;



            //COLUMN B-C
            //הכנסת כותרות של פריטים ל2 עמודות ראשונות
            for (int i = 0; i < 2; i++)//מתחיל מאחרי הכותרות
            {
                for (int j = 0; j < WasteTable.Rows.Count; j++)//
                {
                    xlWorkSheet.Cells[9 + j, i + 2] = WasteTable.Rows[j][i].ToString();
                }
            }


            //העתקת גריד ויו PER DAYS
            Dictionary<int, double> DictionaryWHentrence = new Dictionary<int, double>();
            Dictionary<int, double> DictionarySumCuring = new Dictionary<int, double>();
            DictionarySumCuring = WasteObject.BringSumCuring();
            DictionaryWHentrence = WasteObject.BringWarehouseEntrance();//סכומי כניסה למחסן

            int RowsTotalPerDay = 0;//מספר שורה באקסל של הסיכומים


            //column I until End
            //יצירת כותרות עמודות-של הימים
            int Column = 9;//יתחיל מטור 9
            for (int i = 1; i <= WasteObject.Days; i++)//מספר הימים בחודש יגדיר עמודות יתחיל מטור 3 
            {
                xlWorkSheet.Cells[6, Column] = i;//יום 1 יכניס בשורה 6 טור 3
                xlWorkSheet.Cells[7, Column] = "Weight";
                Column++;
            }

            //COLUMN I UNTIL END
            ///הפירוט לרמת הימים
            //מתחיל אחרי הכותרות
            for (int i = 2; i < WasteTable.Columns.Count - 5; i++)
            {
                int SumScrap = 0, j;
                for (j = 0; j < WasteTable.Rows.Count; j++)//לא כולל טוטל ויעדים
                {
                    xlWorkSheet.Cells[9 + j, i + 7] = WasteTable.Rows[j][i].ToString();//I+6=COLUMN G ,J+9=ROW 9 START
                    bool isNumeric = int.TryParse(WasteTable.Rows[j][i].ToString(), out int n);
                    if (isNumeric)
                    {
                        SumScrap += n;
                    }
                }
                if (i == 2)//כותרות סיכומים
                {
                    //B22
                    RowsTotalPerDay = 9 + j;
                    xlWorkSheet.Cells[RowsTotalPerDay, "B"] = "Total Scrap";
                    xlWorkSheet.Cells[RowsTotalPerDay + 1, "B"] = " Warehous Entrances";
                    xlWorkSheet.Cells[RowsTotalPerDay + 2, "B"] = "Total Cured";
                }
                //מילוי שורות 22-24
                ////סכום פסולות יומי
                if (SumScrap != 0)
                    xlWorkSheet.Cells[RowsTotalPerDay, i + 7] = SumScrap;
                ////סכום כניסה למחסן יומי
                if (DictionaryWHentrence.ContainsKey(int.Parse(WasteTable.Columns[i].ColumnName)))//אם היה בכלל הכנסה למחסן באותו יום ספציפי
                {
                    DictionaryWHentrence.TryGetValue(int.Parse(WasteTable.Columns[i].ColumnName), out double value);
                    xlWorkSheet.Cells[RowsTotalPerDay + 1, i + 7] = value;
                }
                //if (i<TableWarehouse.Rows.Count)
                //xlWorkSheet.Cells[RowsTotalPerDay+1, i + 4] = TableWarehouse.Rows[i]["Sum"].ToString();//j+1 שורה מתחת,i+4 start column D
                ////סכום גיפור יומי
                if (DictionarySumCuring.ContainsKey(int.Parse(WasteTable.Columns[i].ColumnName)))
                {
                    DictionarySumCuring.TryGetValue(int.Parse(WasteTable.Columns[i].ColumnName), out double value);
                    xlWorkSheet.Cells[RowsTotalPerDay + 2, i + 7] = value;
                }
            }


            //עיצוב דטה גריד
            int RowsBorder = DgvWasteMonth.RowCount + 11;//-יתחיל בשורה 9 גבולות תא של דטה גריד באקסל
            xlWorkSheet.get_Range("B6" + ":AM" + RowsBorder).Borders.Color = System.Drawing.Color.Black.ToArgb();
            xlWorkSheet.get_Range("B6" + ":AM" + RowsBorder).Font.Bold = true;
            xlWorkSheet.get_Range("D6" + ":AM" + RowsBorder).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("B6" + ":AM" + RowsBorder).Font.Size = 12;
            xlWorkSheet.get_Range("B6" + ":AM" + RowsBorder).RowHeight = 26;
            xlWorkSheet.get_Range("A6:AK7").RowHeight = 40;//גובה שורה של העמודות

            //צביעת שלושת השורות של הסיכומים
            xlWorkSheet.get_Range("B" + RowsTotalPerDay + ":AM" + (RowsTotalPerDay + 2)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.get_Range("B" + RowsTotalPerDay + ":AM" + (RowsTotalPerDay + 2)).Font.Bold = true;
            xlWorkSheet.get_Range("B" + RowsTotalPerDay + ":AM" + (RowsTotalPerDay + 2)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט


            //הקפאת עמודות
            Excel.Range myCell = (Excel.Range)xlWorkSheet.Cells[8, "D"];
            myCell.Activate();
            myCell.Application.ActiveWindow.FreezePanes = true;

   
            //D-H COLUMN
            // טורי טוטל יעדים ועמידה ביעדים
            xlWorkSheet.Cells[6, "D"] = "TOTAL";
            xlWorkSheet.get_Range("D6:D7").Merge();
            xlWorkSheet.get_Range("D6:D7").Font.Size = 11;
            xlWorkSheet.get_Range("D6:D7").Font.Bold = true;
            xlWorkSheet.get_Range("D6:D7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("D6:D7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.get_Range("D6:E7").WrapText = true;
            xlWorkSheet.get_Range("D6:D7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.Cells[6, "E"] = "Scrap % On cured Gross Tonnage";
            xlWorkSheet.get_Range("E6:E7").Merge();
            xlWorkSheet.get_Range("E6:E7").Font.Size = 11;
            xlWorkSheet.get_Range("E6:E7").Font.Bold = true;
            xlWorkSheet.get_Range("E6:E7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("E6:E7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.get_Range("E6:E7").WrapText = true;
            xlWorkSheet.get_Range("E6:E7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.Cells[6, "F"] = "Scrap % On cured Neto Tonnage";// 
            xlWorkSheet.get_Range("F6:F7").Merge();
            xlWorkSheet.get_Range("F6:F7").Font.Size = 11;
            xlWorkSheet.get_Range("F6:F7").Font.Bold = true;
            xlWorkSheet.get_Range("F6:F7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("F6:F7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.get_Range("F6:F7").WrapText = true;
            xlWorkSheet.get_Range("F6:F7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.Cells[6, "G"] = "Mtd Scrap% Warehouse"; 
            xlWorkSheet.get_Range("G6:G7").Merge();
            xlWorkSheet.get_Range("G6:G7").Font.Size = 11;
            xlWorkSheet.get_Range("G6:G7").Font.Bold = true;
            xlWorkSheet.get_Range("G6:G7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("G6:G7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.get_Range("G6:G7").WrapText = true;
            xlWorkSheet.get_Range("G6:G7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.get_Range("D6:G" + RowsBorder).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות בולטים להבלטת טוטל ויעדים
            xlWorkSheet.get_Range("D6:G" + RowsBorder).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.Cells[6, "H"] = "Scrap TGT %"; 
            xlWorkSheet.get_Range("H6:H7").Merge();
            xlWorkSheet.get_Range("H6:H7").Font.Size = 11;
            xlWorkSheet.get_Range("H6:H7").Font.Bold = true;
            xlWorkSheet.get_Range("H6:H7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("H6:H7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlWorkSheet.get_Range("H6:H7").WrapText = true;
            xlWorkSheet.get_Range("H6:H7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.get_Range("D6:H" + RowsBorder).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות בולטים להבלטת טוטל ויעדים
            xlWorkSheet.get_Range("D6:H" + RowsBorder).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;

            //צביעת תאים של היעדים
            for (int i = 0; i < DgvWasteMonth.RowCount; i++)//על המילון שנוצר
            {
                string value;
                WasteObject.WasteGreenOrRed.TryGetValue(i, out value);
                if (value == "green")
                    xlWorkSheet.Cells[i + 9, "G"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);

                else if (value == "red")
                    xlWorkSheet.Cells[i + 9, "G"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }


            //column d-h 
            double SumPrecentCured = 0, SumPrecentWH = 0, SumPrecentTarget = 0,SumPrecentNeto=0;
            int r = 0;
            //טוטל כולל עמודה D
            //הכנסת סיכומים (הטורים הכללים)לעמודות הנכונות
            for (int i = 0; i < WasteTable.Rows.Count; i++)
            {
                int counter = 0;
                for (int j = WasteTable.Columns.Count; j > WasteTable.Columns.Count - 5; j--)
                {
                    xlWorkSheet.Cells[i + 9, 8 - counter] = WasteTable.Rows[i][j - 1].ToString();//טור H 8-COUNTER
                    counter++;
                    if (!string.IsNullOrEmpty(WasteTable.Rows[i][j - 1].ToString()))
                    {
                        switch (WasteTable.Columns[j - 1].ColumnName)//סכימה של טורי הסיכומים
                        {
                            case "Scrap % On cured Gross Tonnage":
                                SumPrecentCured += double.Parse(WasteTable.Rows[i][j - 1].ToString().Substring(0, WasteTable.Rows[i][j - 1].ToString().IndexOf("%")));
                                break;

                            case "Scrap % On cured Neto Tonnage":
                                SumPrecentNeto += double.Parse(WasteTable.Rows[i][j - 1].ToString().Substring(0, WasteTable.Rows[i][j - 1].ToString().IndexOf("%")));
                                break;

                            case "Mtd Scrap% Warehouse":
                                SumPrecentWH += double.Parse(WasteTable.Rows[i][j - 1].ToString().Substring(0, WasteTable.Rows[i][j - 1].ToString().IndexOf("%")));
                                break;

                            case "Scrap TGT %":
                                SumPrecentTarget += double.Parse(WasteTable.Rows[i][j - 1].ToString().Substring(0, WasteTable.Rows[i][j - 1].ToString().IndexOf("%")));
                                break;
                        }
                    }
                }
                r = i + 1;
            }
            xlWorkSheet.Cells[9 + r, "E"] = (SumPrecentCured / 100).ToString("P2");
            xlWorkSheet.Cells[9 + r, "F"] = (SumPrecentNeto / 100).ToString("P2");
            xlWorkSheet.Cells[9 + r, "G"] = (SumPrecentWH / 100).ToString("P2");
            xlWorkSheet.Cells[9 + r, "H"] = (SumPrecentTarget / 100).ToString("P2");


            //Rows BottomLine 25-31
            int RowEndTable = RowsBorder + 1;
            RowsBorder++;
            //הכנסת סיכומים
            xlWorkSheet.Cells[RowsBorder, "B"] = "Total Cured-kg:";
            xlWorkSheet.Cells[RowsBorder, "C"] = Convert.ToDecimal(WasteObject.SumProduction).ToString("#,#"); ;
            RowsBorder++;
            xlWorkSheet.Cells[RowsBorder, "B"] = "Total Scrap:";
            xlWorkSheet.Cells[RowsBorder, "C"] = Convert.ToDecimal(WasteObject.SumWaste).ToString("#,#"); ;
            RowsBorder++;
            xlWorkSheet.Cells[RowsBorder, "B"] = "Total Neto:";
            xlWorkSheet.Cells[RowsBorder, "C"] = Convert.ToDecimal(WasteObject.ProductionNeto).ToString("#,#"); ;
            RowsBorder++;
            xlWorkSheet.Cells[RowsBorder, "B"] = "Entry to Warehouse:";
            xlWorkSheet.Cells[RowsBorder, "C"] = Convert.ToDecimal(WasteObject.EntryToWarehouse).ToString("#,#"); ;
            RowsBorder++;
            xlWorkSheet.Cells[RowsBorder, "B"] = "Scrap Target %:";
            xlWorkSheet.Cells[RowsBorder, "C"] = Convert.ToDecimal(WasteObject.ScrapTgtPrecent).ToString() + "%"; ;
            RowsBorder++;
            xlWorkSheet.Cells[RowsBorder, "B"] = "Total Scrap %:";
            xlWorkSheet.Cells[RowsBorder, "C"] = Convert.ToDecimal(WasteObject.ScarpTotalPrecent).ToString("p2");
            if (WasteObject.ScrapTgtPrecent < WasteObject.ScarpTotalPrecent * 100)
                xlWorkSheet.Cells[RowsBorder, "C"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            else
                xlWorkSheet.Cells[RowsBorder, "C"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);


            Excel.Range chartTotalRange = xlWorkSheet.get_Range("B" + RowEndTable, "C" + RowsBorder);

            chartTotalRange.Font.Size = 14;
            chartTotalRange.Font.Bold = true;
            chartTotalRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot;//גבולות בולטים להבלטת טוטל ויעדים
            chartTotalRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot;//גבולות בולטים להבלטת טוטל ויעדים
            chartTotalRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot;
            //xlWorkSheet.get_Range("E1:F1").ColumnWidth = 35;


            xlexcel.Visible = true;
            Cursor.Current = Cursors.Default;
           
            releaseObject(xlexcel);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);
        }


        private void Data_Desc_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void DgvWasteMonth_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //צביעת תאים של היעדים
            for (int i = 0; i < DgvWasteMonth.RowCount; i++)//על המילון שנוצר
            {
                string value;
                WasteObject.WasteGreenOrRed.TryGetValue(i, out value);
                if (value == "green")
                    DgvWasteMonth.Rows[i].Cells["Mtd Scrap% Warehouse"].Style.BackColor = Color.YellowGreen;

                else if (value == "red")
                    DgvWasteMonth.Rows[i].Cells["Mtd Scrap% Warehouse"].Style.BackColor = Color.Red;
            }


            //צביעת תאים שעודכנו כבר
            string DayColumn = "";
            for (int i = 0; i < DataOfCellTable.Rows.Count; i++)
            {
                DayColumn = DataOfCellTable.Rows[i]["MSQTDT"].ToString().Substring(6, 2);
                for (int j = 0; j < DgvWasteMonth.RowCount; j++)
                {
                    if (DataOfCellTable.Rows[i]["MSMKT"].ToString() == DgvWasteMonth.Rows[j].Cells["Catalog Number"].Value.ToString())//לולאה על דטה גריד עד שמוצאים את השורה של המקט
                    {
                        if (DayColumn.Substring(0, 1) == "0") DayColumn = DayColumn.TrimStart('0');
                        DgvWasteMonth.Rows[j].Cells[DayColumn].Style.BackColor = Color.Turquoise;
                    }
                }

            }

            //צביעת תא שעודכן עכשיו
            if (CellUpdated && SaveData)
            {
                if (string.IsNullOrEmpty(DataGrid_Desc.Rows[0].Cells[0].EditedFormattedValue as string))//אם טבלה ריקה זה אומר שלא היה עדכון
                    DgvWasteMonth.Rows[eRowIndex].Cells[eColumnIndex].Style.BackColor = Color.White;
                else
                    DgvWasteMonth.Rows[eRowIndex].Cells[eColumnIndex].Style.BackColor = Color.Turquoise;
                SaveData = false;//יחסום צביעת תא אם לא שמרנו נתונים                
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void DataGrid_Desc_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            LeftOver += int.Parse(DataGrid_Desc.Rows[e.Row.Index].Cells["כמות"].Value.ToString());
            lbl_totalLeftover.Text = LeftOver.ToString();
            if(LeftOver>0)  lbl_totalLeftover.ForeColor = System.Drawing.Color.Black;
            else if (LeftOver < 0) lbl_totalLeftover.ForeColor = System.Drawing.Color.Red;

        }
    }
}
