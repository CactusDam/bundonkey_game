using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        Image[] walkleft = new Image[3];
        Image[] walkright = new Image[3];
        Image[] walkup = new Image[4];
        Image[] walkdown = new Image[4];
        string[] INV = new string[10];
        int INVSPACE = 9;
        public Form1()
        {
            KeyDown += new KeyEventHandler(Form1_KeyDown);
            InitializeComponent();
            walkleft[0] = Properties.Resources.GBwalkleft1;
            walkleft[1] = Properties.Resources.GBwalkleft2;
            walkleft[2] = Properties.Resources.GBwalkleft3;
            walkright[0] = Properties.Resources.GBwalkright1;
            walkright[1] = Properties.Resources.GBwalkright2;
            walkright[2] = Properties.Resources.GBwalkright3;
            walkup[0] = Properties.Resources.GBwalkup1;
            walkup[1] = Properties.Resources.GBwalkup2;
            walkup[2] = Properties.Resources.GBwalkup1;
            walkup[3] = Properties.Resources.GBwalkup3;
            walkdown[0] = Properties.Resources.GBwalkdown1;
            walkdown[1] = Properties.Resources.GBwalkdown2;
            walkdown[2] = Properties.Resources.GBwalkdown1;
            walkdown[3] = Properties.Resources.GBwalkdown3;
            Excel(expath, 1);
        }
        
        string expath = @"C:/Users/conti/Desktop/Old_Finals/final-2018-f/gamefile.xlsx";
        int sheet = 1;
        int bg = 0;
        //bad boy is 0
        int pb1spriteid = 0;

        Application excel = new Excel.Application();
        Workbook workbook;
        Worksheet worksheet;
        public void Excel(string expath, int sheet)
        {
            this.expath = expath;
            this.sheet = sheet;
            workbook = excel.Workbooks.Open(expath);
            worksheet = workbook.Worksheets[sheet];
            if(worksheet.Cells[2,12].Value=="RIVER")
            {
                BackgroundImage = Properties.Resources.rivergif;
                bg = 1;
            }
            if (worksheet.Cells[2,12].Value == "VILLAGE")
            {
                BackgroundImage = Properties.Resources.villagegif;
                bg = 0;
            }
        }



        public void Save()
        {
            workbook.Save();
        }


        public string whatweget(int row, int col)
        {
            string[] names = new string[15]; //actually starts at 2
            names[0] = "EXODIME"; names[1] = "SHYNIK"; names[2] = "ERPAN";
            names[3] = "CICLID"; names[4] = "BONAPARTE"; names[5] = "PANDER";
            names[6] = "PAWS"; names[7] = "CLUTZ"; names[8] = "AIDIDAN";
            names[9] = "DEPORY"; names[10] = "FOURITE"; names[11] = "RIVIL";
            names[12] = "ANTIDAN"; names[13] = "CHAWUNGA"; names[14] = "MCAFEE";

            string[] row1 = new string[11]; //starts at 1
            row1[0] = "NAME"; row1[1] = "ID"; row1[2] = "LOC"; row1[3] = "SPRITENUM";
            row1[4] = "STATUS"; row1[5] = "POW"; row1[6] = "HEALTH"; row1[7] = "AGI";
            row1[8] = "LVL"; row1[9] = "DESC"; row1[10] = "DIALOGUE";

            string get=worksheet.Cells[row, col].Value;

            return "";
        }

        public string readit(int cellnum, int rownum)
        {
           // string temp = "";

            if(worksheet.Cells[cellnum, rownum].Value2!=null)
            {
                return worksheet.Cells[cellnum, rownum];
            }
            else { return ""; }
        }

        //int width = 132; int height = 130;
        int x = 100; int y = 100; int speed=5;
        int walk = 0;
        void Form1_KeyDown(object sen, KeyEventArgs e)
        {
            

            if (bg == 0)
            {
                pictureBox1.Image = Properties.Resources.Exodime;
                pb1spriteid = 1;
            }
            if(bg==1)
            {
                pictureBox1.Image = Properties.Resources.Bad_Boy_1_;
                pb1spriteid = 0;
            }

            if (e.KeyCode==Keys.A)
            {
                if (walk < 3)
                {
                    pictureBox2.Image = walkleft[walk];
                    walk++;
                }
                else { pictureBox2.Image = walkleft[0]; walk = 0; } 
                y-=speed;

            }
            else if(e.KeyCode==Keys.W)
            {
                if (walk < 4)
                {
                    pictureBox2.Image = walkup[walk];
                    walk++;
                }
                else { pictureBox2.Image = walkup[0]; walk = 0; }
                x -=speed;
            }
            else if(e.KeyCode==Keys.D)
            {
                if (walk < 3)
                {
                    pictureBox2.Image = walkright[walk];
                    walk++;
                }
                else { pictureBox2.Image = walkright[0]; walk = 0; }
                y +=speed;
            }
            else if(e.KeyCode==Keys.S)
            {
                if (walk < 4)
                {
                    pictureBox2.Image = walkdown[walk];
                    walk++;
                }
                else { pictureBox2.Image = walkdown[0]; walk = 0; }
                x +=speed;
            }

            else if(e.KeyCode==Keys.R)
            {
                worksheet.Cells[2, 12].Value = "AAAA";
                Save();
            }

            else if(e.KeyCode==Keys.Escape)
            {
                workbook.Close();
                this.Close();
            }

            else if(e.KeyCode==Keys.G)
            {
                textBox2.Visible = false;
            }

            else if(e.KeyCode==Keys.I)
            {
                INVENTORY.Visible = true;
                INVENTORY.Top = 0;
                if(INV[0]!=null)
                {
                    label1.Text = INV[0];
                    label1.Visible = true;
                }
            }

            if(e.KeyCode==Keys.X)
            {
                INVENTORY.Visible = false;
                pictureBox4.Visible = false;
                textBox1.Visible = false;
                textBox2.Visible = false;
                label1.Visible = false;
            }


            if (e.KeyCode == Keys.Space && (pictureBox1.Top - pictureBox2.Top < 200 || pictureBox1.Left - pictureBox2.Left < 200))
            {
                //MessageBox.Show("SPACE");
                timer1.Enabled=false;
                timer1.Start();
                timer1.Interval = (10000);
                
                if (pb1spriteid==0)
                {
                    textBox1.Visible = true;
                    textBox1.Top = 0;
                    textBox1.Text = worksheet.Cells[10, 10].value + "";
                }
                if(pb1spriteid==1)
                {
                    textBox1.Visible = true;
                    textBox1.Top = 0;
                    textBox1.Text = worksheet.Cells[2, 10].Value+" ";
                }
            }

            if (isithit(x, y))
            {
                if(pictureBox3.Bounds.IntersectsWith(pictureBox2.Bounds))
                {
                    pictureBox4.Visible = true;
                    if (e.KeyCode == Keys.Space)
                    {
                        if (bg == 0)
                        {
                            worksheet.Cells[2, 12].Value = "VILLAGE";
                            Save();
                        }
                        if (bg == 1)
                        {
                            worksheet.Cells[2, 12].Value = "RIVER";
                            Save();
                        }
                        pictureBox4.Visible = false;
                    }
                    else
                        pictureBox4.Visible = false;
                }
                if (y > pictureBox1.Right)
                {
                    y -= 20;
                }
                if (y < pictureBox1.Right)
                {
                    y += 20;
                }
                if (x > pictureBox1.Bottom)
                {
                    x = 150;
                }

            }

            if (e.KeyCode == Keys.D && pictureBox2.Right >= 900 && bg == 0)
            {
                //to river
                y = 100;
                pictureBox2.Left = y;
                BackgroundImage = Properties.Resources.rivergif;
                bg = 1;
            }
            if (e.KeyCode == Keys.A && pictureBox2.Left <= 40 && bg == 1)
            {
                //to village
                BackgroundImage = Properties.Resources.villagegif;
                y = 800;
                pictureBox2.Left = y;
                bg = 0;
            }

            else
                pictureBox2.Top = x;
                pictureBox2.Left = y;
            
        }

        

        public bool isithit(int x, int y)
        {
            //NPC

            //MessageBox.Show(x+" "+pictureBox1.Top+" "+pictureBox1.Bottom);
            if (pictureBox1.Bounds.IntersectsWith(pictureBox2.Bounds))
            {
                return true;
            }

            //ITEMS

            if (pictureBox2.Bounds.IntersectsWith(POTIONS.Bounds)&&POTIONS.Visible==true)
            {
                if (INVSPACE >= 0)
                {
                    POTIONS.Visible = false;
                    if (POTIONS.Image == Properties.Resources.redpotion)
                    {
                        INV[9 - INVSPACE] = "RED POTION";
                    }
                    else
                    {
                        INV[9 - INVSPACE] = "POTION";
                    }
                    INVSPACE--;
                }
                
                else
                {
                    MessageBox.Show("INVENTORY FULL");
                }
               // MessageBox.Show("" + INV[0]);
                return false;
            }

            //SAVE
            if (pictureBox3.Bounds.IntersectsWith(pictureBox2.Bounds))
            {
                pictureBox4.Visible = true;
                pictureBox4.Top = 0;
                pictureBox4.Left = 0;
                return true;
            }
            else
            return false; 
        }

 //       int tme = 10;

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            textBox1.Visible = false;
            textBox1.Top = -222;
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    if(bg==0)
        //    {
        //        worksheet.Cells[2, 12].Value = "VILLAGE";
        //        Save();
        //    }
        //    else
        //    {
        //        worksheet.Cells[2, 12].Value = "RIVER";
        //        Save();
        //    }
        //    pictureBox4.Visible = false;
        //    button1.Visible = false;
        //    button2.Visible = false;
        //    pictureBox4.Top = 470;
        //}

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    pictureBox4.Visible = false;
        //    button1.Visible = false;
        //    button2.Visible = false;
        //    pictureBox4.Top =470;

        //}



    }


}
