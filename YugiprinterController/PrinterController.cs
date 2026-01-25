using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Drawing.Printing;
using System.IO;

namespace YugiprinterController
{
    public partial class PrinterController: Form
    {
        //in kh sat: width = 81.9899944414f, height = 243.764172336f;
        //in sat: width = 81.9899944414f, height = 243.764172336f;
        public float x = -35f, y = -50f, width = 81.9899944414f, height = 243.764172336f;

        private string pageSize = "0", pageSide = "0", fileFormat = "0", exportFileName = "TestNew.docx";
        private string[] tempArray;

        int horizontalIndex = 0, verticalIndex = 0;
        float horizontalValue = 180, verticalValue = 250;
        string userPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public PrinterController()
        {
            InitializeComponent();
        }

        private void PrinterController_Load(object sender, EventArgs e)
        {
            var settingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                "PRD Team", "YugipriterSetting", "appSetting.txt");
            var appSettings = AppSetting.Load(settingsFile);

            string selectedDownFolder = appSettings.SelectedFolderPath;
            bool printClose = appSettings.PrintCloseToCard;

            if (printClose)
            {
                x = -35f;
                y = -50f;
                width = 81.9899944414f;
                height = 243.764172336f;
                horizontalValue = 167.5f;
                verticalValue = 243.8f;
            }
            else
            {
                x = -35f;
                y = -50f;
                width = 81.9899944414f;
                height = 243.764172336f;
                horizontalValue = 180;
                verticalValue = 250;
            }

            horizontalIndex = 0;
            verticalIndex = 0;
            bool a3PageColum = false, sideBackground = false;

            // 1. Xác định đường dẫn file
            string myDocPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(myDocPath, "PRD Team", "YugipriterSetting");
            string filePath = Path.Combine(myDocPath, "PRD Team", "YugipriterSetting", "settingDeckString.txt");

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                if (!File.Exists(filePath))
                {
                    File.WriteAllText(filePath, string.Empty);
                }

                if (File.Exists(filePath))
                {
                    tempArray = File.ReadAllLines(filePath);

                    if (tempArray.Length == 0)
                    {
                        MessageBox.Show("File dữ liệu trống!", "Thông báo");
                        Application.Exit();
                        //return;
                    }
                }
                else
                {
                    MessageBox.Show("Không tìm thấy file tại: " + filePath, "Lỗi file");
                    Application.Exit();
                    //return;
                }

                Document document = new Document();
                document.LoadFromFile("Doc1.docx");


                foreach (Section section in document.Sections)
                {
                    if (pageSize == "0")
                    {
                        section.PageSetup.PageSize = PageSize.A4;
                    }
                    else if (pageSize == "1")
                    {
                        section.PageSetup.PageSize = PageSize.A3;
                    }
                }

                for (int i = 0; i < tempArray.Length; i++)
                {
                    Section section = document.Sections[0];

                    if (pageSize == "0")
                    {
                        if (pageSide == "0")
                        {
                            if ((i == 9 || i == 18 || i == 27 || i == 36 || i == 45 || i == 54 || i == 63 || i == 72 || i == 81) && i != 0)
                            {
                                Paragraph paragraph = section.Paragraphs[0];
                                paragraph.AppendBreak(BreakType.PageBreak);
                                verticalIndex = 0;
                            }

                            DocPicture picture = section.Paragraphs[0].AppendPicture(Image.FromFile(selectedDownFolder + "/" + tempArray[i] + ".jpg"));
                            picture.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                            picture.VerticalPosition = y + (verticalValue * verticalIndex);

                            horizontalIndex += 1;

                            if (horizontalIndex > 2)
                            {
                                horizontalIndex = 0;
                                verticalIndex += 1;
                            }

                            picture.Width = width;
                            picture.Height = height;

                            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                        }
                        else if (pageSide == "1")
                        {
                            if ((i == 9 || i == 18 || i == 27 || i == 36 || i == 45 || i == 54 || i == 63 || i == 72 || i == 81) && i != 0)
                            {
                                Paragraph paragraph = section.Paragraphs[0];
                                paragraph.AppendBreak(BreakType.PageBreak);
                                verticalIndex = 0;
                                sideBackground = true;
                            }

                            if (sideBackground)
                            {
                                for (int i2 = 0; i2 < 9; i2++)
                                {
                                    DocPicture picture2 = section.Paragraphs[0].AppendPicture(Image.FromFile($"{Application.StartupPath}\\{"YgoBackCard.png"}"));
                                    picture2.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                                    picture2.VerticalPosition = y + (verticalValue * verticalIndex);

                                    horizontalIndex += 1;

                                    if (horizontalIndex > 2)
                                    {
                                        horizontalIndex = 0;
                                        verticalIndex += 1;
                                    }

                                    picture2.Width = width;
                                    picture2.Height = height;

                                    picture2.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                                }

                                sideBackground = false;
                                Paragraph paragraph2 = section.Paragraphs[0];
                                paragraph2.AppendBreak(BreakType.PageBreak);
                                horizontalIndex = 0;
                                verticalIndex = 0;
                            }

                            DocPicture picture = section.Paragraphs[0].AppendPicture(Image.FromFile(selectedDownFolder + "/" + tempArray[i] + ".jpg"));
                            picture.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                            picture.VerticalPosition = y + (verticalValue * verticalIndex);

                            horizontalIndex += 1;

                            if (horizontalIndex > 2)
                            {
                                horizontalIndex = 0;
                                verticalIndex += 1;
                            }

                            picture.Width = width;
                            picture.Height = height;

                            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                        }
                    }
                    else if (pageSize == "1")
                    {
                        if (pageSide == "0")
                        {
                            if ((i == 18 || i == 36 || i == 54 || i == 72) && i != 0)
                            {
                                Paragraph paragraph = section.Paragraphs[0];
                                paragraph.AppendBreak(BreakType.PageBreak);
                                horizontalIndex = 0;
                                verticalIndex = 0;
                                a3PageColum = false;
                            }

                            DocPicture picture = section.Paragraphs[0].AppendPicture(Image.FromFile(selectedDownFolder + "/" + tempArray[i] + ".jpg"));
                            picture.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                            picture.VerticalPosition = y + 70 + (verticalValue * verticalIndex);

                            if (a3PageColum == true)
                            {
                                picture.Rotation = 90;
                                picture.HorizontalPosition += 32;
                                picture.VerticalPosition = y + 8 + ((verticalValue - 75) * verticalIndex);
                            }

                            if (a3PageColum == false)
                            {
                                horizontalIndex += 1;
                            }
                            else if (a3PageColum == true)
                            {
                                verticalIndex += 1;
                            }

                            if (horizontalIndex == 3 && verticalIndex < 4 && a3PageColum == false)
                            {
                                horizontalIndex = 0;
                                verticalIndex += 1;
                            }

                            if (verticalIndex == 4 && a3PageColum == false)
                            {
                                a3PageColum = true;
                                verticalIndex = 0;
                                horizontalIndex = 3;
                            }

                            picture.Width = width;
                            picture.Height = height;

                            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                        }
                        else if (pageSide == "1")
                        {
                            if ((i == 18 || i == 36 || i == 54 || i == 72) && i != 0)
                            {
                                Paragraph paragraph = section.Paragraphs[0];
                                paragraph.AppendBreak(BreakType.PageBreak);
                                horizontalIndex = 0;
                                verticalIndex = 0;
                                a3PageColum = false;
                                sideBackground = true;
                            }

                            if (sideBackground)
                            {
                                for (int i2 = 0; i2 < 18; i2++)
                                {
                                    DocPicture picture2 = section.Paragraphs[0].AppendPicture(Image.FromFile($"{Application.StartupPath}\\{"YgoBackCard.png"}"));
                                    picture2.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                                    picture2.VerticalPosition = y + 70 + (verticalValue * verticalIndex);

                                    if (a3PageColum == true)
                                    {
                                        picture2.Rotation = 90;
                                        picture2.HorizontalPosition += 32;
                                        picture2.VerticalPosition = y + 8 + ((verticalValue - 75) * verticalIndex);
                                    }

                                    if (a3PageColum == false)
                                    {
                                        horizontalIndex += 1;
                                    }
                                    else if (a3PageColum == true)
                                    {
                                        verticalIndex += 1;
                                    }

                                    if (horizontalIndex == 3 && verticalIndex < 4 && a3PageColum == false)
                                    {
                                        horizontalIndex = 0;
                                        verticalIndex += 1;
                                    }

                                    if (verticalIndex == 4 && a3PageColum == false)
                                    {
                                        a3PageColum = true;
                                        verticalIndex = 0;
                                        horizontalIndex = 3;
                                    }

                                    picture2.Width = width;
                                    picture2.Height = height;

                                    picture2.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                                }

                                sideBackground = false;
                                a3PageColum = false;
                                Paragraph paragraph2 = section.Paragraphs[0];
                                paragraph2.AppendBreak(BreakType.PageBreak);
                                horizontalIndex = 0;
                                verticalIndex = 0;
                            }

                            DocPicture picture = section.Paragraphs[0].AppendPicture(Image.FromFile(selectedDownFolder + "/" + tempArray[i] + ".jpg"));
                            picture.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                            picture.VerticalPosition = y + 70 + (verticalValue * verticalIndex);

                            if (a3PageColum == true)
                            {
                                picture.Rotation = 90;
                                picture.HorizontalPosition += 32;
                                picture.VerticalPosition = y + 8 + ((verticalValue - 75) * verticalIndex);
                            }

                            if (a3PageColum == false)
                            {
                                horizontalIndex += 1;
                            }
                            else if (a3PageColum == true)
                            {
                                verticalIndex += 1;
                            }

                            if (horizontalIndex == 3 && verticalIndex < 4 && a3PageColum == false)
                            {
                                horizontalIndex = 0;
                                verticalIndex += 1;
                            }

                            if (verticalIndex == 4 && a3PageColum == false)
                            {
                                a3PageColum = true;
                                verticalIndex = 0;
                                horizontalIndex = 3;
                            }

                            picture.Width = width;
                            picture.Height = height;

                            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                        }
                    }
                }

                if (pageSize == "0")
                {
                    if (pageSide == "1")
                    {
                        Section sectionLastPage = document.Sections[0];

                        Paragraph paragraphLastPage = sectionLastPage.Paragraphs[0];
                        paragraphLastPage.AppendBreak(BreakType.PageBreak);
                        verticalIndex = 0;
                        horizontalIndex = 0;
                        sideBackground = true;

                        for (int i2 = 0; i2 < 9; i2++)
                        {
                            DocPicture pictureLastPage = sectionLastPage.Paragraphs[0].AppendPicture(Image.FromFile($"{Application.StartupPath}\\{"YgoBackCard.png"}"));
                            pictureLastPage.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                            pictureLastPage.VerticalPosition = y + (verticalValue * verticalIndex);

                            horizontalIndex += 1;

                            if (horizontalIndex > 2)
                            {
                                horizontalIndex = 0;
                                verticalIndex += 1;
                            }

                            pictureLastPage.Width = width;
                            pictureLastPage.Height = height;

                            pictureLastPage.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                        }

                        sideBackground = false;
                        horizontalIndex = 0;
                        verticalIndex = 0;
                    }
                }
                else if (pageSize == "1")
                {
                    if (pageSide == "1")
                    {
                        Section sectionLastPage = document.Sections[0];

                        Paragraph paragraphLastPage = sectionLastPage.Paragraphs[0];
                        paragraphLastPage.AppendBreak(BreakType.PageBreak);
                        verticalIndex = 0;
                        horizontalIndex = 0;
                        sideBackground = true;
                        a3PageColum = false;

                        for (int i2 = 0; i2 < 18; i2++)
                        {
                            DocPicture pictureLastPage = sectionLastPage.Paragraphs[0].AppendPicture(Image.FromFile($"{Application.StartupPath}\\{"YgoBackCard.png"}"));
                            pictureLastPage.HorizontalPosition = x + (horizontalValue * horizontalIndex);
                            pictureLastPage.VerticalPosition = y + 70 + (verticalValue * verticalIndex);

                            if (a3PageColum == true)
                            {
                                pictureLastPage.Rotation = 90;
                                pictureLastPage.HorizontalPosition += 32;
                                pictureLastPage.VerticalPosition = y + 8 + ((verticalValue - 75) * verticalIndex);
                            }

                            if (a3PageColum == false)
                            {
                                horizontalIndex += 1;
                            }
                            else if (a3PageColum == true)
                            {
                                verticalIndex += 1;
                            }

                            if (horizontalIndex == 3 && verticalIndex < 4 && a3PageColum == false)
                            {
                                horizontalIndex = 0;
                                verticalIndex += 1;
                            }

                            if (verticalIndex == 4 && a3PageColum == false)
                            {
                                a3PageColum = true;
                                verticalIndex = 0;
                                horizontalIndex = 3;
                            }

                            pictureLastPage.Width = width;
                            pictureLastPage.Height = height;

                            pictureLastPage.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                        }

                        sideBackground = false;
                        a3PageColum = false;
                        horizontalIndex = 0;
                        verticalIndex = 0;
                    }
                }

                if (fileFormat == "0")
                {
                    //exportFileName = ExportFileNameTextBox.Text + ".docx";
                    document.SaveToFile(userPath + "\\" + exportFileName, FileFormat.Docx);
                }
                else if (fileFormat == "1")
                {
                    //exportFileName = ExportFileNameTextBox.Text + ".pdf";
                    document.SaveToFile(userPath + "\\" + exportFileName, FileFormat.PDF);
                }

                MessageBox.Show("Xuất file thành công!\nĐường dẫn file tại " + userPath + "\\" + exportFileName, "Thành công");
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                Application.Exit();
            }
        }
    }
}
