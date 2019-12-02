using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WasteManagement
{
    class Waste
    {
        public string CatalogNumber { get; set; }// זיהוי קטלוג
        public string  DescriptionCatalog { get; set; }//תיאור קטלוג
        public string DateByS400Start { get; set; }//תאריך התחלה בs400 לדוגמא 901101\
        public string DateByS400End { get; set; }//תאריך סיום לפי s400
        public int AmountWaste { get; set; }// כמות טוטלית ממק"ט מסוים (שורה)
        public int AmountPerDay { get; set; }//כמות של יום ספציפי
        public double SumProduction { get; set; }//כמות כוללת של ייצור עבור כל החודש
        public double SumWaste { get; set; }//כמות כוללת של פסולת עבור כל החודש
        public double ProductionNeto { get; set; }//כמות נטו של ייצור(כמות כוללת פחות כמות פסולת)
        public double EntryToWarehouse { get; set; }//כמות של כמה נכנס למחסן באותו חודש
        public string MonthNumber { get; set; }// חודש נבחר לדוח
        public int Year { get; set; }//שנה נבחרת לדוח
        public int YearS400 { get; set; }//שנה לפי s400
        public int Days { get; set; }//כמה ימים בחודש הנבחר
        public double ScrapTgtPrecent { get; set; }//אחוז כולל מטרת פסולת
        public double ScarpTotalPrecent { get; set; }//אחוז פסולת כולל בפועל
        public bool ThisMonth { get; set; }//החודש הנוכחי או רטרו אחורה,נועד בשביל השאילתות שסופרות האם לקחת יום אחורה או לא
        public string Yesterday { get; set; }
        DBService DBS = new DBService();
        DataTable DTgeneral = new DataTable();//טבלה שאוספת את כל הנתונים מהחודש הנדרש
        DataTable DtToShow ;//טבלת ההצגה הסופית לפי פורמט דו"ח
        DataTable DtCatalogNumber;//טבלה שמכילה את כל המק"טים והתיאור שלהם-נועד לשים בכל תחילת חודש את השמות גם אם אין פירוט
        DataTable DtOfRowTiresDisallowd = new DataTable();//שורה שתווסף לכלל השורות-סכום צמיגים ספציפיים שנפסלו
        public Dictionary<int, string> WasteGreenOrRed { get; set; }//ירוק או אדום מבחינת יעדים
        public Dictionary<int, string> CodeAndDescWaste { get; set; }//קוד פסולת ותיאורו

        public Waste()
        {
            ScrapTgtPrecent = 0;//מאפס אחוזי פסולת
            ScarpTotalPrecent = 0;
            Yesterday = DateTime.Today.AddDays(-1).Day.ToString();
            if (DateTime.Today.AddDays(-1).Day < 10)
                Yesterday = "0" + DateTime.Today.AddDays(-1).Day.ToString();
        }


        public DataTable GetWasteList(string MonthNumber,int Year,int Days,bool ThisMonth)
        {
            //שליפת נתוני פסולות
            this.MonthNumber = MonthNumber;
            this.Year = Year;
            YearS400 = Year - 1928;//השנה המוזרה לפי s400
            this.Days = Days;//כמה ימים בחודש הנבחר
            DateByS400Start = YearS400.ToString() + MonthNumber + "01";//בניית תאריך התחלה
            DateByS400End=YearS400.ToString() + MonthNumber + Days.ToString();//בניית תאריך סיום
            this.ThisMonth = ThisMonth;
            DtToShow = new DataTable();//כל פעם ייצור טבלת ייצוג חדשה
            string qry = $@"SELECT b.TPROD,idesc, date('20'||(substring((ttdte),1,2)-72)||'-'||substring(ttdte,3,2)||'-'||substring(ttdte,5,2)) as date  , sum(b.TQTY) as sum  ,day('20'||(substring((ttdte),1,2)-72)||'-'||substring(ttdte,3,2)||'-'||substring(ttdte,5,2)) as day
                          FROM  BPCSFV30.ithl01 as b left join BPCSFV30.IIML01 as c on iprod=tprod                          WHERE TTYPE ='PX' and TWHS ='SC' and TTDTE between {DateByS400Start} and {DateByS400End}
                          GROUP BY b.TPROD,idesc, b.TTDTE                              ORDER BY date";
        
            // MONTH (('20'||(substring((ttdte),1,2)-72)||'-'||substring(ttdte,3,2)||'-'||substring(ttdte,5,2)))={MonthNumber}  and YEAR (('20'||(substring((ttdte),1,2)-72)||'-'||substring(ttdte,3,2)||'-'||substring(ttdte,5,2))) ={Year} 
            DTgeneral = DBS.executeSelectQueryNoParam(qry);//שליפת רשימת כל הפריטים
            string qryCatalog = $@"SELECT MSMKT15 as TPROD ,idesc 
                                 FROM msk.mshb as m left join BPCSFV30.IIML01 as c on  c.iprod=m.MSMKT15 
                                 WHERE MSSUG ='1'";
            DtCatalogNumber = new DataTable();
            DtCatalogNumber = DBS.executeSelectQueryNoParam(qryCatalog);
            BringRowTiresDisallowd();
            DesignTable();
          
            return DtToShow;

        }

        /// <summary>
        ///שלוש שורות שנוספו  לפי בקשתו של אלחנן-פסולות של הייצור עצמו לפני שריטה שוקלת,אמור לצאת כמויות דומות למה שריטה מקלידה. 
        ///p1-sc-ct350,p2-sc-rd010,p3-sc-ct375
        /// </summary>
        private void BringRowTiresDisallowd()
        {
            //string qry = $@"SELECT date('20'||(substring((C.SSVGDT),1,2)-72)||'-'||substring(C.SSVGDT,3,2)||'-'||substring(C.SSVGDT,5,2)) as date,int(sum(D.MKG)) as sum, day('20'||(substring((C.SSVGDT),1,2)-72)||'-'||substring(C.SSVGDT,3,2)||'-'||substring(C.SSVGDT,5,2)) as day
            //            FROM ALIQUAL.CTSVGP C,BPCSFALI.IIMNL99 I,aliqual.TPGAMP P,ALLTAB.COSTP D, DAVIDM.DATENEW K ,alltab.t1601p 
            //            WHERE substring(C.SMAKAT, 1, 8) = I.INPROD AND C.SPGAM1 = P.CDPGAM AND C.SSVGDT  between {DateByS400Start} and {YearS400.ToString() + MonthNumber + Yesterday}  AND substring(C.SMAKAT,1,8) =D.CATNUM and ssivug='פ' and I.INPROD= T16PRD and T16ADF='*'
            //            GROUP BY  C.SSVGDT";//שאילתה של דוד מרקוביץ-השתנה וכבר לא רלוונטי
            string qry = $@"SELECT   date('20'||(substring((ttdte),1,2)-72)||'-'||substring(ttdte,3,2)||'-'||substring(ttdte,5,2)) as date,int(SUM(ICSCP1)) as sum ,day('20'||(substring((ttdte),1,2)-72)||'-'||substring(ttdte,3,2)||'-'||substring(ttdte,5,2)) as day,ttype
                           FROM BPCSFV30.ITH  as b  inner join BPCSFV30.CICL01 as c on b.TPROD=c.ICPROD
                           WHERE TTDTE BETWEEN {DateByS400Start} AND {YearS400.ToString() + MonthNumber + Yesterday} and TTYPE in('P3','P2','P1')
                           GROUP BY b.TTDTE,TTYPE 
                           ORDER BY TTYPE";
            DtOfRowTiresDisallowd = DBS.executeSelectQueryNoParam(qry);
        }

        /// <summary>
        /// עיצוב טבלה לפי דו"ח מתבקש
        /// </summary>
        private void DesignTable()
        {
            //יצירת עמודות
            DtToShow.Columns.Add("Catalog Number");//עמודת שם חומר
            DtToShow.Columns.Add("description");//עמודת שם חומר


            for (int i = 1; i <= Days; i++)
                {
                    //if (i < 10)
                    //    DtToShow.Columns.Add("0" + i.ToString());//כל יום בתור עמודה
                    //else
                    DtToShow.Columns.Add(i.ToString());
                }
            DtToShow.Columns.Add("Total");
            DtToShow.Columns.Add("Scrap % On cured Gross Tonnage");
            DtToShow.Columns.Add("Scrap % On cured Neto Tonnage");
            DtToShow.Columns.Add("MTD Scrap% WAREHOUSE  TONNAGE");//אחוז פסולת כללית מתוך כלל הייצור
            DtToShow.Columns.Add("Scrap TGT %");//מטרת אחוז פסולת

            //טבלה של מספרים קטלוגים ותיאור הפריט,יצור את שורות הטבלת הצגה
            //יצירת שורות
            DataView view = new DataView(DtCatalogNumber);
            DataTable TableCatlaogNames = view.ToTable(true, "TPROD","idesc");//טבלה מסייעת עבור שמות
            
            //הכנסת cured tires scrap-actually instead of rita typing
            DataRow[] row = TableCatlaogNames.Select("TPROD='SC-CT350'");
            int index = TableCatlaogNames.Rows.IndexOf(row[0]);//index of rita typing
            DataRow CuredScrapActually = TableCatlaogNames.NewRow();
            CuredScrapActually["TPROD"] = "_SC-CT350";
            CuredScrapActually["idesc"] = "Cured Tires Scraps-Actually";
            TableCatlaogNames.Rows.InsertAt(CuredScrapActually, index);//מכניס שורה לפני השורה של cured tires scrap
          
            for (int i = 0; i < TableCatlaogNames.Rows.Count; i++)
            {
                DtToShow.Rows.Add(TableCatlaogNames.Rows[i]["TPROD"]);
                DtToShow.Rows[i]["description"] = TableCatlaogNames.Rows[i]["idesc"];          
            }

            string WhichDay;
            //מילוי השורות לפי ימים
            for (int i = 0; i < DtToShow.Rows.Count; i++)//ממלא כל תיאור פריט בנפרד
            {
                double Total = 0;
                for (int j = 0; j < DTgeneral.Rows.Count; j++)//מחפש תיאורי פריט תואמים מטבלה כוללת
                {
                    if (DtToShow.Rows[i][0].ToString() == DTgeneral.Rows[j]["TPROD"].ToString())
                    {
                        WhichDay = DTgeneral.Rows[j]["day"].ToString();//לאיזה עמודת יום להכניס את הנתון
                        double Sum = double.Parse(DTgeneral.Rows[j]["sum"].ToString());//סכום הפסולת
                        DtToShow.Rows[i][WhichDay]=Sum.ToString();//מכניס לרשומה הרלוונטית
                        Total += Sum;//מוסיף לטוטל
                    }
                }
                DtToShow.Rows[i]["Total"] = Convert.ToDecimal(Total).ToString("#,#");
            }

            double SumTiresDisallowd = 0;
            //מילוי שורת cured tires scrap-actually 
            for (int i = 0; i < DtOfRowTiresDisallowd.Rows.Count; i++)
            {
                WhichDay = DtOfRowTiresDisallowd.Rows[i]["day"].ToString();
                DtToShow.Rows[index][WhichDay] = double.Parse(DtOfRowTiresDisallowd.Rows[i]["sum"].ToString());
                SumTiresDisallowd += double.Parse(DtOfRowTiresDisallowd.Rows[i]["sum"].ToString());
            }
            DtToShow.Rows[index]["Total"] = Convert.ToDecimal(SumTiresDisallowd).ToString("#,#");

            SumOfEntryWarehouse();//סכום כניסה למחסן
            GetSumProduction();//סכום גיפור כולל-צד ימין למטה
            FillTargets();//מילוי יעדים
            GetSumWaste();////סכום כולל פסולת


        }

        /// <summary>
        /// שליפת יעדים עבור חומרים מסוימים
        /// </summary>
        private void FillTargets()
        {
            WasteGreenOrRed = new Dictionary<int, string>();//לדעת אם צריך לצבוע תא באדום או ירוק
            DataTable TargetsTable = new DataTable();
            string qry = $@"SELECT *
                           FROM BPCSFALI.QLVL01
                           WHERE QVYEAR={Year.ToString().Substring(2, 2)} and QVMNTH ={MonthNumber} and QVDAY=1";
            TargetsTable = DBS.executeSelectQueryNoParam(qry);
            double ScrapPrecentOfWarehouse = 0;
            double ScrapPrecentOFcured = 0;
            //הכנסת יעדים לכל שורה שיש בה יעד ולפי השדה שנקבע לה 
            for (int i = 0; i < DtToShow.Rows.Count; i++)
            {
                if (TargetsTable.Rows.Count!=0)
                {
                    if (DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "_SC-CT350")//Cured Tires left overs        
                        DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCRP"] + "%";
                    else if (DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "SC-FC010")//Rubberized Textile left overs 
                        DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCRT"] + "%";
                    else if (DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "SC-FC050") //Rubberized Steel left overs   
                        DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCRS"] + "%";
                    else if (DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "FB-40002")//Rubber Compound Grade - B     
                        DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCRR"] + "%";
                    else if (DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "SC-GT300")//Green Tires left overs        
                        DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCRG"] + "%";
                    else if (DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "SC-BE200")
                        DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCRB"] + "%";  //bead scrap                
                    //else if(DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "SC-BL100")
                    //    DtToShow.Rows[i]["Scrap TGT %"] = TargetsTable.Rows[0]["QVSCBL"] + "%";  //bladers                

                }


                //חישוב אחוזי פסולת מתוך כמות כוללת
                if (!string.IsNullOrEmpty(DtToShow.Rows[i]["Scrap TGT %"].ToString() as string) && EntryToWarehouse!=0 && SumProduction!=0)
                {
                    //if(DtToShow.Rows[i]["TPROD"].ToString().Trim()== "SC-FC050") ScrapPrecent= double.Parse(DtToShow.Rows[i]["Total"].ToString()+ double.Parse(DtToShow.Rows[i]["Total"].ToString()
                    ScrapPrecentOfWarehouse = (double.Parse(DtToShow.Rows[i]["Total"].ToString()) / EntryToWarehouse);//אחוז פסולת המוצר מתוך כניסה למחסן 
                    ScrapPrecentOFcured = (double.Parse(DtToShow.Rows[i]["Total"].ToString()) / SumProduction);//אחוז פסולת המוצר מתוך ייצור  כולל
                    DtToShow.Rows[i]["MTD Scrap% WAREHOUSE  TONNAGE"] = ScrapPrecentOfWarehouse.ToString("p3");//2 מספרים אחרי הנקודה באחוזים
                    DtToShow.Rows[i]["Scrap % On cured Gross Tonnage"] = ScrapPrecentOFcured.ToString("p3");
                    ScrapTgtPrecent += double.Parse(DtToShow.Rows[i]["Scrap TGT %"].ToString().Remove(DtToShow.Rows[i]["Scrap TGT %"].ToString().LastIndexOf('%')));//חישוב אחוז מטרה כולל של פסולת,כל פעם מוסיף לטוטל 
                    //ScarpTotalPrecent += ScrapPrecent;//כל פעם מוסיף את אחוז הפסולת עד שיש כמות כוללת
                   
                        if (ScrapPrecentOfWarehouse*100 < double.Parse(DtToShow.Rows[i]["Scrap TGT %"].ToString().Remove(DtToShow.Rows[i]["Scrap TGT %"].ToString().LastIndexOf('%'))))//אם כמות הפסולת קטנה מהיעד ייצבע בירוק, אחרת באדום
                        WasteGreenOrRed.Add(i, "green");
                    else
                        WasteGreenOrRed.Add(i, "red");
                }
                ScrapPrecentOfWarehouse = 0;
            }          

        }

        /// <summary>
        /// סכום של כלל הגיפורים של אותו חודש
        /// </summary>
        /// <returns></returns>
        public void GetSumProduction()
        {
            DataTable dataTable = new DataTable();
            string qry="";
            
            if (ThisMonth)//אם הדוח עבור החודש הזה לוקחים יום אחורה ,אם חודשים אחורה אז לוקחים את כל הימים
            {
                qry = $@"SELECT sum( TPCS * TWGHT ) AS SUM
                           FROM   BPCSFV30 .fltl01 
                           WHERE TDEPT in (161,162) and ttdte between {DateByS400Start} and {YearS400.ToString() + MonthNumber + Yesterday} ";//זה דדיווחי עבודה כלליים
            }
            else
            {
                qry = $@"SELECT sum( TPCS * TWGHT ) AS SUM
                           FROM   BPCSFV30 .fltl01 
                           WHERE TDEPT in (161,162) and ttdte between {DateByS400Start} and {DateByS400End} ";
            }
       
            dataTable = DBS.executeSelectQueryNoParam(qry);
            SumProduction = 0;
            if (!string.IsNullOrEmpty(dataTable.Rows[0]["SUM"].ToString() as string))
                SumProduction = double.Parse(dataTable.Rows[0]["SUM"].ToString());
        }

        /// <summary>
        ///כמות כוללת נטו ו סכום כולל של כל הפסולות עבור כל החודש
        /// </summary>
        public void GetSumWaste()
        {
            double ScrapPrecentOfNeto = 0;
            SumWaste = 0;
            for (int i = 0; i < DtToShow.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(DtToShow.Rows[i]["Total"].ToString() as string) && (!string.IsNullOrEmpty(DtToShow.Rows[i]["Scrap TGT %"].ToString()) || DtToShow.Rows[i]["Catalog Number"].ToString().Trim() == "SC-FC055"))
                    SumWaste += double.Parse(DtToShow.Rows[i]["Total"].ToString());
            }
           
            //לא צריך, שינו את הנוסחה לחישוב כולל-תשמור להמשך
            ProductionNeto = SumProduction - SumWaste;//כמות נטו של ייצור
            if (ProductionNeto != 0)
            {
                //בהמשך לנסות לפתור את זה בלי הלולאה
                //אחוזי פסולות נטו-אפשר פה כי זה הגופר פחות הפסולות ורק פה אני מקבל את סכום הפסולות
                for (int i = 0; i < DtToShow.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(DtToShow.Rows[i]["Scrap TGT %"].ToString() as string)&& !string.IsNullOrEmpty(DtToShow.Rows[i]["Total"].ToString() as string))
                    {
                        ScrapPrecentOfNeto = (double.Parse(DtToShow.Rows[i]["Total"].ToString()) / (ProductionNeto));//אחוז פסולת המוצר מתוך נטוווו הייצור כלומר פחות הפסולת
                        DtToShow.Rows[i]["Scrap % On cured Neto Tonnage"] = ScrapPrecentOfNeto.ToString("p3");
                    }
                }
            }
            if (EntryToWarehouse!=0)
            ScarpTotalPrecent = SumWaste / EntryToWarehouse;

        }

        /// <summary>
        /// משיג פרטי קודי פסולת של החודש הנוכחי
        /// </summary>
        /// <returns></returns>
        public DataTable GetCellsData()
        {
            DataTable GetCellsDataTable = new DataTable();
            string qry = $@"SELECT *
                            FROM MSK.MSVQTP
                            WHERE MSQTDT between {Year + MonthNumber}01 and {Year + MonthNumber + Days}";
            GetCellsDataTable = DBS.executeSelectQueryNoParam(qry);
            return GetCellsDataTable;
        }


        /// <summary>
        /// מחיקת עדכון של תא
        /// </summary>
        public void DeleteCellUpdate()
        {
            string qry = $@"delete 
                          FROM MSK.MSVQTP
                          WHERE MSMKT='{CatalogNumber}' and MSQTDT between {Year + MonthNumber}01 and {Year + MonthNumber + Days}";
            DBS.executeInsertQuery(qry);
        }

        /// <summary>
        /// רשימת קודי פסולות עבור הפירוט-ימולא בקומבובוקס
        /// </summary>
        public DataTable GetWasteCodes()
        {
            DataTable GetWasteCodesTable = new DataTable();
            string qry = $@"SELECT   MSSKOD as code, MSDEC as desc         
                          FROM msk. MSSVGP";
            GetWasteCodesTable = DBS.executeSelectQueryNoParam(qry);
            CodeAndDescWaste = new Dictionary<int, string>();//הוספה למילון קוד פסולת ותיאורו
            for (int i = 0; i < GetWasteCodesTable.Rows.Count; i++)
            {
                CodeAndDescWaste.Add(int.Parse(GetWasteCodesTable.Rows[i]["code"].ToString()), GetWasteCodesTable.Rows[i]["desc"].ToString());
            }
            return GetWasteCodesTable;
        }

        /// <summary>
        /// מביא סכום כניסה למחסןו
        /// </summary>
        public void SumOfEntryWarehouse()
        {
            DataTable dataTable = new DataTable();
            string qry = "";
            if (ThisMonth)//אם הדוח עבור החודש הזה לוקחים יום אחורה ,אם חודשים אחורה אז לוקחים את כל הימים
            {
                qry = $@"SELECT round(sum(TQTY*ICSCP1),1) as Sum
                           FROM BPCSFV30.ITHL01 left join BPCSFV30.CICL01 on TPROD=ICPROD and ICFAC='F1'                           WHERE substring(TTYPE,1,1) ='T' and twhs in ('HD','GM') and THTOWH ='AL' and TTDTE between {DateByS400Start} and {YearS400.ToString() + MonthNumber +Yesterday} ";//עד יום אחורה
            }
            else
            {
                qry = $@"SELECT round(sum(TQTY*ICSCP1),1) as Sum
                           FROM BPCSFV30.ITHL01 left join BPCSFV30.CICL01 on TPROD=ICPROD and ICFAC='F1'                           WHERE substring(TTYPE,1,1) ='T' and twhs in ('HD','GM') and THTOWH ='AL' and TTDTE between {DateByS400Start} and {DateByS400End} " ;
            }
            dataTable = DBS.executeSelectQueryNoParam(qry);
            if (!string.IsNullOrEmpty(dataTable.Rows[0]["Sum"].ToString() as string))
                EntryToWarehouse = double.Parse(dataTable.Rows[0]["Sum"].ToString());

        }


        /// <summary>
        ///  Warehouse entrance טוטל ימים לאקסל עבור 
        /// </summary>
        internal Dictionary<int,double> BringWarehouseEntrance()
        {
            DataTable TableWarehouse = new DataTable();
            Dictionary<int, double> DictionaryWHentrence = new Dictionary<int, double>();
            string qry = $@" SELECT round(sum(TQTY*ICSCP1)/1000,1) as Sum,substring(ttdte,5,2) as Day
                           FROM BPCSFV30.ITHL01 left join BPCSFV30.CICL01 on TPROD=ICPROD and ICFAC='F1'                           WHERE substring(TTYPE,1,1) ='T' and twhs in ('HD','GM') and THTOWH ='AL' and TTDTE between {DateByS400Start} and {DateByS400End} 
                           GROUP BY ttdte";
            TableWarehouse = DBS.executeSelectQueryNoParam(qry);
            for (int i = 0; i < TableWarehouse.Rows.Count; i++)
            {
                DictionaryWHentrence.Add(int.Parse(TableWarehouse.Rows[i]["Day"].ToString()), double.Parse(TableWarehouse.Rows[i]["Sum"].ToString()));
            }
            return DictionaryWHentrence;

        }

        /// <summary>
        /// טוטל של ימים עבור כמה גופרו סה"כ
        /// </summary>
        /// <returns></returns>
        internal Dictionary<int, double> BringSumCuring()
        {
            DataTable TableSumCuring = new DataTable();
            Dictionary<int, double> DictionarySumCuring = new Dictionary<int, double>();
            string qry = $@" select substring(ttdte,5,2) as Day, round(sum(TPCS*TWGHT)/1000,2) as Sum  
                             from BPCSFALI.FLTQV 
                             where TTDTE between {DateByS400Start} and {DateByS400End} and TDEPT in (161,162) group by TTDTE";
            TableSumCuring = DBS.executeSelectQueryNoParam(qry);
            for (int i = 0; i < TableSumCuring.Rows.Count; i++)
            {
                DictionarySumCuring.Add(int.Parse(TableSumCuring.Rows[i]["Day"].ToString()), double.Parse(TableSumCuring.Rows[i]["Sum"].ToString()));
            }
            return DictionarySumCuring;
        }
    }
}
