using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents;
using SD = Spire.Doc.Documents;
using System.Text.RegularExpressions;
using System.IO;
using TextBox = Spire.Doc.Fields.TextBox;
using System.Drawing.Imaging;
using System.Net;

namespace PDFCore
{
    class Program
    {
        public Dictionary<string, string> answers = new Dictionary<string, string>();
        public Dictionary<int, string> headings = new Dictionary<int, string>();
        public int HeadingsCounter = 0;
        public string LoadFilePath = string.Empty;
        public string AnswerFilePath = string.Empty;
        public string parameters = "";

        static void Main(string[] args)
        {
            //Spire.License.LicenseProvider.SetLicenseKey("cNY4LA4awojvAQAc3lE9K05POKItvvpuz9UuEPmAHJJnlKXCgctCnBleTPUTfIG5FJzTzqA/YSOUC32vtcOLU/ActYKge1TrlTO/xqWPfNBAdlkVpM/zsYX2C1srea7KYSSUbGWDNqS8cezA5c/PcaLx98W60IJ4o8xS7zBWyByjEAM9YLx4bxz19gd38APu2AndnSUeRSUCJQsMtnj+o1HFbavchlu7ti1Lifzpwef9mUBcAXi8VCbfaCN9IML9DC6rDKCmCtBxTF4ey6IAbVU/r174hv5Stc1+GRLPxrZXnGG+L6vw4EnKIJddeZyuHuudHDCRRIg72PkdcnnfB47cCXSygzofvRLCtrmaXQgdsIM0Z0dh76E1+NCGwbSWEma4KAtgA1D4e1E5u1BB0nLDhANGGCLLmbRtaFNe8lZifdlfFPI/SBrSgYPhr4Vql2Km7HodfIzQWkduAMMRBQwsmDN4Qz+oNfvVJUATKl59h/LposnY6WFVuFJtFKDTT5Acwe1XgsmIpfAx0N96C3fmAqkxb6Gcz0q1mtfkjo4ECAVshPkxVaZKRzg333UZJggaeSJ+cjzMZXJ4x+KpbEKc/GiNH8esxYKG8b1IWvszJLSOgMrW8NtPpA+xc1+Hjh0ODV8OfAJNw31aqMmsUmD2FkZQHxppCOoUdtMcsAm4Lo0J1V11kD9Szd7x/4dLtZDo2MQVgGZGi+5IY7Ve9R8YK8Km17MDyc6HHx0R0VXV7GIu56UTb7Z8znfxiAlKWUn/r4I+NvlVYR+W0HAtKfdAUJUQ7/kOxHulWm4Ke+ZHM5vWkAR6ERXw2kirjMVDScxb0w2SpHJK+JPKsxkug0s+CYmsMqlTsKG1RuEyl3VPnF4TiYEjkx0goL5M7IuHYwTw2PL3Kvu4X3qGgCHdb/MDRyBXUO8UGzhiTZ3mMMwAfYisBZ0jgj58Yr7sGKFy1e9o+d1CuG0vDMDMMXuv4apdKHVJF4BviVJDEPNpgAUQK6BZ0waexZVw5O4XTsBo51fcsKDaPyMSnY5iGJz/3Cge2krQ2mz59ZYzh+cNbx/graOilPdFIUE2tmdT/EMZpHgFQpRfnaJ9W0UGO/nR+7pfYnFDDB2ZGZk4B+HtjyAk9J7vQbCEQ/bVcGuCbxjdDhGgfS7U70VSyP5b5X7aZkeIU/r8cJLLfK6QWrDXYjeIBSoQto0S6AF+CATglztCPQ1PsYaXHR2Ab74ANUWHhaX9U0US59cf/7zWJV1PzCycW90r0l09KxU3adpO7RbPIvmpUtVO25dc8Vdwdh3cKASbanMXHTBttAR85PLKP9RHPEIN/+FvwPIo3QLDGNLh3s9OOoZvAT6UPNuDzUWadFnv+/pxb02fKCxUNycYez2/n4lTT8c4sbUJT81UrRG2uCHqMXdMyzTikV8eztSKA9yyyYnfZlEcIqbPPY0aBrmHuEHH6STGd3qNA/u6Rs9sCfIOcb3seGh6ywXWn1WYl86kXr3boGX+0g18gPgSxfg=");
            Program obj = new Program();
            // string param = "";
            if (args.Length > 0)
            {
                if (args.Length == 1)
                {
                    obj.parameters = args[0];
                }
                else
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        obj.parameters += args[i] + "&";

                    }
                    obj.parameters = obj.parameters.TrimEnd('&');
                }


            }
            else
            {
                Console.WriteLine("Expecting parameters");
                System.Environment.Exit(0);
            }
            obj.parameters = obj.parameters.Replace("%", " ");
            // Console.WriteLine(obj.parameters);


            //Program obj = new Program();
            obj.initComp();
            //obj.disp();
            Console.WriteLine("Analyzing files...");

            obj.extractImages();

            Console.WriteLine("Image extraction done");

            obj.imageTags();

            Console.WriteLine("Image tags insertion done");

            obj.readContents();

            Console.WriteLine("Reading docx contents");

            obj.LoadContent();

            Console.WriteLine("Extracting questions");

            obj.CleanContent();

            Console.WriteLine("Cleaning questions");

            obj.TableOutput();

            Console.WriteLine("Creating ECAT output");

            Console.WriteLine("Done");
            //obj.CreateDoc();

        }

        public void TableOutput()
        {
            Document document = new Document();
            Section section = document.AddSection();
            string correctOptions = "";
            string marksIfTrue = "1";
            string marksIfFalse = "0";
            string isMandatory = "FALSE";
            string questionType = "WORD";
            string createdBy = "PDF Auto";
            string isHidden = "FALSE";
            string currQuestionType = string.Empty;
            int questionNumber = 1;
            string courses = "CBSE";
            string subjects = "Mathematics";
            string tags = "CBSE, Mathematics, Class VI";
            string correctAnswer = "";
            string globalExplanation = "";
            string currentLine = "";
            int tableCount = -1;
            int rowCount = 1;
            Boolean questionStatus = false;
            int lower = 1;
            int upper = 1;
            Boolean tokenstatus = false;
            string skey = "";

            string FILE_NAME = Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-3.txt";


            if (System.IO.File.Exists(FILE_NAME) == true)
            {
                System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);

                while (objReader.Peek() != -1)
                {
                    currentLine = objReader.ReadLine();

                    if (currentLine.Trim().Length < 1)
                    {
                        continue;
                    }

                    string first1 = currentLine.Split('.').First();

                    if (first1.All(char.IsDigit))
                    {
                        try
                        {
                            questionNumber = Convert.ToInt32(first1);
                        }
                        catch (Exception ex)
                        {

                            // Console.WriteLine(ex.Message);
                        }

                    }

                    if (currentLine.StartsWith("[") && currentLine.Contains("UNIT"))
                    {
                        skey = currentLine.Replace("[", "").Replace("]", "");
                        if (answers.ContainsKey(skey))
                        {
                            correctAnswer = answers[skey];
                        }
                        continue;
                    }

                    if (currentLine.Contains("[MCQ]"))
                    {
                        questionType = "MCQ";
                        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
                        {
                            lower = numValue;
                        }
                        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
                        {
                            upper = numValue2;
                        }
                        continue;

                    }
                    else if (currentLine.Contains("[TF]"))
                    {
                        questionType = "TF";
                        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
                        {
                            lower = numValue;
                        }
                        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
                        {
                            upper = numValue2;
                        }
                        continue;
                    }
                    else if (currentLine.Contains("[FIB]"))
                    {
                        questionType = "FIB";

                        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
                        {
                            lower = numValue;
                        }
                        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
                        {
                            upper = numValue2;
                        }
                        continue;
                    }


                    if (questionNumber >= lower && questionNumber <= upper)
                    {
                        currQuestionType = questionType;
                    }
                    else
                    {
                        currQuestionType = "WORD";
                    }

                    //string first = currentLine.Split('.').First();
                    //if(first.All(char.IsDigit))
                    //{
                    //    questionStatus = true;
                    //    continue;
                    //}
                    if (currentLine.Contains("<question>"))
                    {
                        questionStatus = true;
                        tokenstatus = true;
                        continue;
                    }

                    if (questionStatus == true)
                    {
                        section.AddParagraph();
                        Table table = section.AddTable(true);

                        // table.ResetCells(2, 2);
                        rowCount = 1;
                        table.ResetCells(11, 2);
                        table.Rows[2].Cells[1].SplitCell(2, 1);
                        table[1, 0].AddParagraph().AppendText("Question Type");
                        table[1, 1].AddParagraph().AppendText(currQuestionType);
                        table[2, 0].AddParagraph().AppendText("Marks");
                        table[2, 1].AddParagraph().AppendText(marksIfTrue);
                        table[2, 2].AddParagraph().AppendText(marksIfFalse);
                        table[3, 0].AddParagraph().AppendText("isMandatory");
                        table[3, 1].AddParagraph().AppendText(isMandatory);
                        table[4, 0].AddParagraph().AppendText("isHidden");
                        table[4, 1].AddParagraph().AppendText(isHidden);
                        table[5, 0].AddParagraph().AppendText("correctAnswer");
                        table[5, 1].AddParagraph().AppendText(correctAnswer);
                        table[6, 0].AddParagraph().AppendText("globalExplanation");
                        table[6, 1].AddParagraph().AppendText(globalExplanation);
                        table[7, 0].AddParagraph().AppendText("createdBy");
                        table[7, 1].AddParagraph().AppendText(createdBy);
                        table[8, 0].AddParagraph().AppendText("subjects");
                        table[8, 1].AddParagraph().AppendText(subjects);
                        table[9, 0].AddParagraph().AppendText("courses");
                        table[9, 1].AddParagraph().AppendText(courses);
                        table[10, 0].AddParagraph().AppendText("tags");
                        table[10, 1].AddParagraph().AppendText(tags);
                        //table[10, 0].AddParagraph().AppendText();
                        //table[10, 1].AddParagraph().AppendText();

                        //table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                        questionStatus = false;
                        tableCount += 1;
                    }

                    if (tableCount >= 0)
                    {
                        Table ctable = document.Sections[0].Tables[tableCount] as Spire.Doc.Table;
                        ctable.ApplyHorizontalMerge(0, 0, 1);
                        //ctable.Rows[6].Cells[1].SplitCell(2, 1);
                        ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                        if (currQuestionType == "MCQ")
                        {
                            if (currentLine.StartsWith("(a)") || currentLine.StartsWith("(A)"))// && (questionType !="SA" || questionType != "LA" || questionType != "FILEUPLOAD" || questionType != "FILEUPLOAD-D"))
                            {
                                currentLine = currentLine.Replace("(a)", "").Trim();
                                currentLine = currentLine.Replace("(A)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(2, 1);
                                ctable.Rows[rowCount].Cells[2].SplitCell(2, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);

                                if (correctAnswer.Trim() == "(A)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);


                            }
                            else if (currentLine.StartsWith("(b)") || currentLine.StartsWith("(B)"))
                            {
                                currentLine = currentLine.Replace("(b)", "").Trim();
                                currentLine = currentLine.Replace("(B)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(B)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(c)") || currentLine.StartsWith("(C)"))
                            {
                                currentLine = currentLine.Replace("(c)", "").Trim();
                                currentLine = currentLine.Replace("(C)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(C)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(d)") || currentLine.StartsWith("(D)"))
                            {
                                currentLine = currentLine.Replace("(d)", "").Trim();
                                currentLine = currentLine.Replace("(D)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(D)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(e)") || currentLine.StartsWith("(E)"))
                            {
                                currentLine = currentLine.Replace("(e)", "").Trim();
                                currentLine = currentLine.Replace("(E)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(E)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(f)") || currentLine.StartsWith("(F)"))
                            {
                                currentLine = currentLine.Replace("(f)", "").Trim();
                                currentLine = currentLine.Replace("(F)", "").Trim();
                                TableRow row = ctable.AddRow();

                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(2, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(F)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else
                            {
                                string first = currentLine.Split('.').First();
                                if (first.All(char.IsDigit))
                                {
                                    currentLine = currentLine.Replace(first + ".", "");
                                }
                                ctable[0, 0].AddParagraph().AppendText(currentLine);

                            }


                        }
                        else if (currQuestionType == "TF")
                        {
                            TableRow row = ctable.AddRow();
                            rowCount += 1;
                            ctable.Rows.Insert(rowCount, row);
                            ctable[rowCount, 0].AddParagraph().AppendText("True");
                            if (correctAnswer == "T")
                            {
                                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                            }

                            TableRow row2 = ctable.AddRow();
                            rowCount += 1;
                            ctable.Rows.Insert(rowCount, row2);
                            ctable[rowCount, 0].AddParagraph().AppendText("False");
                            if (correctAnswer == "F")
                            {
                                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                            }

                            string first = currentLine.Split('.').First();
                            if (first.All(char.IsDigit))
                            {
                                currentLine = currentLine.Replace(first + ".", "");
                            }
                            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
                            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
                            ctable[0, 0].AddParagraph().AppendText(currentLine);


                        }
                        else if (currQuestionType == "FIB")
                        {
                            string[] tokens;

                            if (correctAnswer.Contains(";"))
                            {
                                tokens = correctAnswer.Split(';');
                            }
                            else if (correctAnswer.Contains("(a)"))
                            {
                                string answer = correctAnswer.Trim().Replace("(a)", "").Replace("(b)", "-").Replace("(c)", "-").Replace("(d)", "-").Replace("(e)", "-");
                                tokens = answer.Split('-');
                            }
                            else
                            {
                                if (correctAnswer.Contains(", "))
                                {
                                    var ans = correctAnswer.Replace(", ", "-");
                                    tokens = ans.Split('-');
                                }
                                else
                                {
                                    tokens = new string[] { correctAnswer };
                                }

                            }

                            if (tokenstatus == true)
                            {
                                int tok = 1;
                                foreach (string item in tokens)
                                {
                                    TableRow row = ctable.AddRow();
                                    rowCount += 1;
                                    ctable.Rows.Insert(rowCount, row);
                                    ctable[rowCount, 0].AddParagraph().AppendText("token" + tok.ToString());
                                    ctable[rowCount, 1].AddParagraph().AppendText(item);
                                    tok += 1;
                                }
                                tokenstatus = false;
                            }

                            string first = currentLine.Split('.').First();
                            if (first.All(char.IsDigit))
                            {
                                currentLine = currentLine.Replace(first + ".", "");
                            }

                            if (currentLine.Trim().Length > 0)
                            {
                                if (!currentLine.Contains("token"))
                                {

                                    if (currentLine.EndsWith("."))
                                    {
                                        currentLine = currentLine.TrimEnd('.') + " [token1].";
                                    }
                                    else
                                    {
                                        currentLine = currentLine + " [token1].";
                                    }
                                }
                                ctable[0, 0].AddParagraph().AppendText(currentLine);
                            }

                            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
                            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');


                        }
                        else
                        {
                            string first = currentLine.Split('.').First();
                            if (first.All(char.IsDigit))
                            {
                                currentLine = currentLine.Replace(first + ".", "");
                            }
                            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
                            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
                            ctable[0, 0].AddParagraph().AppendText(currentLine);
                        }

                    }


                }
                objReader.Dispose();
            }



            //for (int i = 0; i < contentBox.Lines.Count() - 1; i++)
            //{
            //    currentLine = contentBox.Lines[i].ToString();

            //    if(currentLine.Trim().Length <1)
            //    {
            //        continue;
            //    }

            //    string first1 = currentLine.Split('.').First();

            //    if (first1.All(char.IsDigit))
            //    {
            //        try
            //        {
            //            questionNumber = Convert.ToInt32(first1);
            //        }
            //        catch (Exception ex)
            //        {

            //           // Console.WriteLine(ex.Message);
            //        }

            //    }

            //    if(currentLine.StartsWith("[") && currentLine.Contains("UNIT"))
            //    {
            //        skey = currentLine.Replace("[", "").Replace("]", "");
            //        if (answers.ContainsKey(skey))
            //        {
            //            correctAnswer = answers[skey];
            //        }
            //        continue;
            //    }

            //    if (currentLine.Contains("[MCQ]"))
            //    {
            //        questionType = "MCQ";
            //        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
            //        {
            //            lower = numValue;
            //        }
            //        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
            //        {
            //            upper = numValue2;
            //        }
            //        continue;

            //    }
            //    else if (currentLine.Contains("[TF]"))
            //    {
            //        questionType = "TF";
            //        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
            //        {
            //            lower = numValue;
            //        }
            //        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
            //        {
            //            upper = numValue2;
            //        }
            //        continue;
            //    }
            //    else if (currentLine.Contains("[FIB]"))
            //    {
            //        questionType = "FIB";

            //        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
            //        {
            //            lower = numValue;
            //        }
            //        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
            //        {
            //            upper = numValue2;
            //        }
            //        continue;
            //    }


            //    if(questionNumber >= lower && questionNumber <= upper)
            //    {
            //        currQuestionType = questionType;
            //    }
            //    else
            //    {
            //        currQuestionType = "WORD";
            //    }

            //    //string first = currentLine.Split('.').First();
            //    //if(first.All(char.IsDigit))
            //    //{
            //    //    questionStatus = true;
            //    //    continue;
            //    //}
            //    if (currentLine.Contains("<question>"))
            //    {
            //        questionStatus = true;
            //        tokenstatus = true;
            //        continue;
            //    }

            //    if (questionStatus == true)
            //    {
            //        section.AddParagraph();
            //        Table table = section.AddTable(true);

            //        // table.ResetCells(2, 2);
            //        rowCount = 1;
            //        table.ResetCells(11, 2);
            //        table.Rows[2].Cells[1].SplitCell(2, 1);
            //        table[1, 0].AddParagraph().AppendText("Question Type");
            //        table[1, 1].AddParagraph().AppendText(currQuestionType);
            //        table[2, 0].AddParagraph().AppendText("Marks");
            //        table[2, 1].AddParagraph().AppendText(marksIfTrue);
            //        table[2, 2].AddParagraph().AppendText(marksIfFalse);
            //        table[3, 0].AddParagraph().AppendText("isMandatory");
            //        table[3, 1].AddParagraph().AppendText(isMandatory);
            //        table[4, 0].AddParagraph().AppendText("isHidden");
            //        table[4, 1].AddParagraph().AppendText(isHidden);
            //        table[5, 0].AddParagraph().AppendText("correctAnswer");
            //        table[5, 1].AddParagraph().AppendText(correctAnswer);
            //        table[6, 0].AddParagraph().AppendText("globalExplanation");
            //        table[6, 1].AddParagraph().AppendText(globalExplanation);
            //        table[7, 0].AddParagraph().AppendText("createdBy");
            //        table[7, 1].AddParagraph().AppendText(createdBy);
            //        table[8, 0].AddParagraph().AppendText("subjects");
            //        table[8, 1].AddParagraph().AppendText(subjects);
            //        table[9, 0].AddParagraph().AppendText("courses");
            //        table[9, 1].AddParagraph().AppendText(courses);
            //        table[10, 0].AddParagraph().AppendText("tags");
            //        table[10, 1].AddParagraph().AppendText(tags);
            //        //table[10, 0].AddParagraph().AppendText();
            //        //table[10, 1].AddParagraph().AppendText();

            //        //table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //        questionStatus = false;
            //        tableCount += 1;
            //    }

            //    if (tableCount >= 0)
            //    {
            //        Table ctable = document.Sections[0].Tables[tableCount] as Spire.Doc.Table;
            //        ctable.ApplyHorizontalMerge(0, 0, 1);
            //        //ctable.Rows[6].Cells[1].SplitCell(2, 1);
            //        ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //        if(currQuestionType == "MCQ")
            //        {
            //            if (currentLine.StartsWith("(a)") || currentLine.StartsWith("(A)"))// && (questionType !="SA" || questionType != "LA" || questionType != "FILEUPLOAD" || questionType != "FILEUPLOAD-D"))
            //            {
            //                currentLine = currentLine.Replace("(a)", "").Trim();
            //                currentLine = currentLine.Replace("(A)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(2,1);
            //                ctable.Rows[rowCount].Cells[2].SplitCell(2, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);

            //                if(correctAnswer.Trim() == "(A)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);


            //            }
            //            else if (currentLine.StartsWith("(b)") || currentLine.StartsWith("(B)"))
            //            {
            //                currentLine = currentLine.Replace("(b)", "").Trim();
            //                currentLine = currentLine.Replace("(B)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(B)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(c)") || currentLine.StartsWith("(C)"))
            //            {
            //                currentLine = currentLine.Replace("(c)", "").Trim();
            //                currentLine = currentLine.Replace("(C)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(C)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(d)") || currentLine.StartsWith("(D)"))
            //            {
            //                currentLine = currentLine.Replace("(d)", "").Trim();
            //                currentLine = currentLine.Replace("(D)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(D)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(e)") || currentLine.StartsWith("(E)"))
            //            {
            //                currentLine = currentLine.Replace("(e)", "").Trim();
            //                currentLine = currentLine.Replace("(E)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(E)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(f)") || currentLine.StartsWith("(F)"))
            //            {
            //                currentLine = currentLine.Replace("(f)", "").Trim();
            //                currentLine = currentLine.Replace("(F)", "").Trim();
            //                TableRow row = ctable.AddRow();

            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(2, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(F)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else
            //            {
            //                string first = currentLine.Split('.').First();
            //                if (first.All(char.IsDigit))
            //                {
            //                    currentLine = currentLine.Replace(first + ".", "");
            //                }
            //                ctable[0, 0].AddParagraph().AppendText(currentLine);

            //            }


            //        }
            //        else if(currQuestionType == "TF")
            //        {
            //            TableRow row = ctable.AddRow();
            //            rowCount += 1;
            //            ctable.Rows.Insert(rowCount, row);
            //            ctable[rowCount, 0].AddParagraph().AppendText("True");
            //            if (correctAnswer == "T")
            //            {
            //                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //            }

            //            TableRow row2 = ctable.AddRow();
            //            rowCount += 1;
            //            ctable.Rows.Insert(rowCount, row2);
            //            ctable[rowCount, 0].AddParagraph().AppendText("False");
            //            if (correctAnswer == "F")
            //            {
            //                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //            }

            //            string first = currentLine.Split('.').First();
            //            if (first.All(char.IsDigit))
            //            {
            //                currentLine = currentLine.Replace(first + ".", "");
            //            }
            //            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
            //            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
            //            ctable[0, 0].AddParagraph().AppendText(currentLine);


            //        }
            //        else if(currQuestionType == "FIB")
            //        {
            //            string[] tokens;

            //            if(correctAnswer.Contains(";"))
            //            {
            //                tokens = correctAnswer.Split(';');
            //            }
            //            else if(correctAnswer.Contains("(a)"))
            //            {
            //                string answer = correctAnswer.Trim().Replace("(a)", "").Replace("(b)", "-").Replace("(c)", "-").Replace("(d)", "-").Replace("(e)", "-");
            //                tokens = answer.Split('-');
            //            }
            //            else
            //            {
            //                if(correctAnswer.Contains(", "))
            //                {
            //                    var ans = correctAnswer.Replace(", ", "-");
            //                    tokens = ans.Split('-');
            //                }
            //                else
            //                {
            //                    tokens =  new string[] { correctAnswer };
            //                }

            //            }

            //            if (tokenstatus == true)
            //            {
            //                int tok = 1;
            //                foreach (string item in tokens)
            //                {
            //                    TableRow row = ctable.AddRow();
            //                    rowCount += 1;
            //                    ctable.Rows.Insert(rowCount, row);
            //                    ctable[rowCount, 0].AddParagraph().AppendText("token" + tok.ToString());
            //                    ctable[rowCount, 1].AddParagraph().AppendText(item);
            //                    tok += 1;
            //                }
            //                tokenstatus = false;
            //            }

            //            string first = currentLine.Split('.').First();
            //            if (first.All(char.IsDigit))
            //            {
            //                currentLine = currentLine.Replace(first + ".", "");
            //            }

            //            if (currentLine.Trim().Length > 0)
            //            {
            //                if (!currentLine.Contains("token"))
            //                {

            //                    if (currentLine.EndsWith("."))
            //                    {
            //                        currentLine = currentLine.TrimEnd('.') + " [token1].";
            //                    }
            //                    else
            //                    {
            //                        currentLine = currentLine + " [token1].";
            //                    }
            //                }
            //                ctable[0, 0].AddParagraph().AppendText(currentLine);
            //            }

            //            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
            //            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');


            //        }
            //        else 
            //        {
            //            string first = currentLine.Split('.').First();
            //            if (first.All(char.IsDigit))
            //            {
            //                currentLine = currentLine.Replace(first + ".", "");
            //            }
            //            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
            //            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
            //            ctable[0, 0].AddParagraph().AppendText(currentLine);
            //        }

            //    }


            //}


            document.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-final.docx", FileFormat.Docx);
            document.Close();

        }

        public void CleanContent()
        {
            //contentBox.Clear();
            string isMandatory = "";
            string questionType = "";
            string questionNumber = "";
            string currentLine = "";
            int tableCount = -1;
            Boolean questionStatus = false;

            string currentQuestionNo;
            int currentQN = 1;
            int qn = 1;
            int currUnit = 1;
            Boolean unitstatus = true;
            int tokenCount = 1;
            string currLine = "";

            string FILE_NAME = Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.txt";

            using (StreamWriter Wr = new StreamWriter(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-3.txt"))
            {

                if (System.IO.File.Exists(FILE_NAME) == true)
                {
                    System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);

                    while (objReader.Peek() != -1)
                    {
                        currentLine = objReader.ReadLine();

                        currentQuestionNo = currentLine.Split('.')[0];
                        if (currentQuestionNo.Trim().All(char.IsDigit))
                        {
                            try
                            {
                                currentQN = Convert.ToInt32(currentQuestionNo);

                                if (currentQN > 400 || currentQN < 1)
                                {
                                    continue;
                                }

                                if (currentQN >= qn)
                                {
                                    qn = currentQN;
                                    // unitstatus = false;
                                }
                                else
                                {
                                    qn = currentQN;
                                    currUnit += 1;
                                    // unitstatus = true;


                                }

                            }
                            catch (Exception)
                            {

                                continue;
                            }


                            Wr.Write(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "-" + currentQN + "]" + Environment.NewLine);
                            //contentBox.AppendText(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "-" + currentQN + "]" + Environment.NewLine);
                            tokenCount = 1;
                        }

                        string pattern = @"[_]{2,}";
                        Regex regex = new Regex(pattern);
                        foreach (Match ItemMatch in regex.Matches(currentLine))
                        {
                            currentLine = currentLine.Replace(ItemMatch.Value, "_");

                        }


                        int freq = currentLine.Count(f => (f == '_'));

                        for (int fi = 1; fi <= freq; fi++)
                        {
                            string repStr = " [token" + tokenCount.ToString() + "] ";
                            var regex_replace = new Regex(Regex.Escape("_"));

                            currentLine = regex_replace.Replace(currentLine, repStr, 1);
                            tokenCount += 1;
                        }

                        //string first = currentLine.Split('.').First();
                        //if(first.All(char.IsDigit))
                        //{
                        //    textBox1.AppendText(first + Environment.NewLine);
                        //}
                        string first = currentLine.Split('.').First();
                        if (first.All(char.IsDigit))
                        {
                            currentLine = "<question>\\n" + currentLine;
                        }

                        currentLine = currentLine.Replace("(a)", "\\n(a)").Replace("(b)", "\\n(b)").Replace("(c)", "\\n(c)").Replace("(d)", "\\n(d)").Replace("(e)", "\\n(e)").Replace("(f)", "\\n(f)");
                        currentLine = currentLine.Replace("(A)", "\\n(A)").Replace("(B)", "\\n(B)").Replace("(C)", "\\n(C)").Replace("(D)", "\\n(D)").Replace("(E)", "\\n(E)").Replace("(F)", "\\n(F)");

                        var result = currentLine.Split(new string[] { "\\n" }, StringSplitOptions.None);
                        foreach (string s in result)
                        {
                            // contentBox.AppendText(s + Environment.NewLine);
                            Wr.Write(s + Environment.NewLine);
                        }






                    }
                    objReader.Dispose();
                }

            }


            //for (int i = 0; i < headingsBox.Lines.Count() - 1; i++)
            //{
            //    currentLine = headingsBox.Lines[i].ToString();

            //    //if (unitstatus == true)
            //    //{
            //    //    contentBox.AppendText(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "]" + Environment.NewLine);
            //    //}

            //    currentQuestionNo = currentLine.Split('.')[0];
            //    if (currentQuestionNo.Trim().All(char.IsDigit))
            //    {
            //        try
            //        {
            //            currentQN = Convert.ToInt32(currentQuestionNo);

            //            if (currentQN > 400 || currentQN < 1)
            //            {
            //                continue;
            //            }

            //            if (currentQN >= qn)
            //            {
            //                qn = currentQN;
            //               // unitstatus = false;
            //            }
            //            else
            //            {
            //                qn = currentQN;
            //                currUnit += 1;
            //               // unitstatus = true;


            //            }

            //        }
            //        catch (Exception)
            //        {

            //            continue;
            //        }



            //        contentBox.AppendText(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "-" + currentQN + "]" + Environment.NewLine);
            //        tokenCount = 1;
            //    }

            //    string pattern = @"[_]{2,}";
            //    Regex regex = new Regex(pattern);
            //    foreach (Match ItemMatch in regex.Matches(currentLine))
            //    {
            //        currentLine = currentLine.Replace(ItemMatch.Value, "_" );

            //    }


            //    int freq = currentLine.Count(f => (f == '_'));

            //    for(int fi =1; fi <= freq; fi++)
            //    {
            //        string repStr = " [token" + tokenCount.ToString() + "] ";
            //        var regex_replace = new Regex(Regex.Escape("_"));

            //        currentLine = regex_replace.Replace(currentLine, repStr, 1);
            //        tokenCount += 1;
            //    }

            //    //string first = currentLine.Split('.').First();
            //    //if(first.All(char.IsDigit))
            //    //{
            //    //    textBox1.AppendText(first + Environment.NewLine);
            //    //}
            //    string first = currentLine.Split('.').First();
            //    if (first.All(char.IsDigit))
            //    {
            //        currentLine = "<question>\\n" + currentLine;
            //    }

            //    currentLine = currentLine.Replace("(a)", "\\n(a)").Replace("(b)", "\\n(b)").Replace("(c)", "\\n(c)").Replace("(d)", "\\n(d)").Replace("(e)", "\\n(e)").Replace("(f)", "\\n(f)");
            //    currentLine = currentLine.Replace("(A)", "\\n(A)").Replace("(B)", "\\n(B)").Replace("(C)", "\\n(C)").Replace("(D)", "\\n(D)").Replace("(E)", "\\n(E)").Replace("(F)", "\\n(F)");

            //    var result = currentLine.Split(new string[] { "\\n" }, StringSplitOptions.None);
            //    foreach (string s in result)
            //    {
            //        contentBox.AppendText(s + Environment.NewLine);
            //    }

            //}


        }

        public void LoadContent()
        {
            string fileName1 = Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.docx";
            //string fileName1 = @"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Input\Questions11.docx";


            Boolean isBold = false;

            Document doc = new Document();
            doc.LoadFromFile(fileName1);
            using (StreamWriter Wr = new StreamWriter(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.txt"))
            {
                foreach (Section sec in doc.Sections)
                {
                    foreach (Paragraph para in sec.Paragraphs)
                    {

                        isBold = false;
                        foreach (DocumentObject docobj in para.ChildObjects)
                        {
                            if (docobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange text = docobj as TextRange;

                                if (text.CharacterFormat.Bold)
                                {
                                    isBold = true;
                                }
                                else
                                {
                                    isBold = false;
                                }
                                if (text.CharacterFormat.UnderlineStyle == UnderlineStyle.Single)
                                {
                                    text.Text = "_";
                                }
                            }
                        }

                        if (isBold)
                        {
                            // allHeadingsBox.AppendText(para.Text + Environment.NewLine);

                            if (para.Text.Contains("out of the four options") || para.Text.Contains("out of four options") || para.Text.Contains("only one of the four options") || para.Text.Contains("out of the given four options"))
                            {
                                // headingsBox.AppendText("[MCQ]");
                                Wr.Write("[MCQ]");
                                string input = para.Text;
                                // Split on one or more non-digit characters.
                                string[] numbers = Regex.Split(input, @"\D+");
                                foreach (string value in numbers)
                                {
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        //int i = int.Parse(value);
                                        // headingsBox.AppendText("-" + value);
                                        Wr.Write("-" + value);

                                    }
                                }
                                Wr.WriteLine();
                                //headingsBox.AppendText(Environment.NewLine);

                            }
                            else if (para.Text.Contains("true or false") || para.Text.Contains("true (T) or false (F)") || para.Text.Contains("T or F") || para.Text.Contains("(T) or (F)"))
                            {
                                //headingsBox.AppendText("[TF]");
                                Wr.Write("[TF]");
                                string input = para.Text;
                                // Split on one or more non-digit characters.
                                string[] numbers = Regex.Split(input, @"\D+");
                                foreach (string value in numbers)
                                {
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        //int i = int.Parse(value);
                                        // headingsBox.AppendText("-" + value);
                                        Wr.Write("-" + value);

                                    }
                                }
                                // headingsBox.AppendText(Environment.NewLine);
                                Wr.WriteLine();
                            }
                            else if (para.Text.Contains("fill in the blanks"))
                            {
                                //headingsBox.AppendText("[FIB]");
                                Wr.WriteLine("[FIB]");
                                string input = para.Text;
                                // Split on one or more non-digit characters.
                                string[] numbers = Regex.Split(input, @"\D+");
                                foreach (string value in numbers)
                                {
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        //int i = int.Parse(value);
                                        // headingsBox.AppendText("-" + value);
                                        Wr.Write("-" + value);

                                    }
                                }
                                //headingsBox.AppendText(Environment.NewLine);
                                Wr.WriteLine();
                            }
                            //else
                            //{
                            //    headingsBox.AppendText("[WORD]" + Environment.NewLine);
                            //}

                        }
                        else
                        {
                            // headingsBox.AppendText(para.ListText + para.Text + Environment.NewLine);
                            Wr.Write(para.ListText + para.Text);
                            Wr.WriteLine();
                        }

                    }
                }
                Wr.Close();
            }


        }

        public void readContents()
        {
            Document doc = new Document();

            doc.LoadFromFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-1.docx");

            // doc.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Exemplar.docx");

            Document doc2 = new Document();

            Section s2 = doc2.AddSection();
            s2.PageSetup.PageSize = PageSize.A4;

            doc2.Sections[0].PageSetup.Margins.Top = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Bottom = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Left = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Right = 10.9f;

            Boolean qstatus = false;
            int optstatus = 0;
            Boolean headStatus = false;
            Boolean isBold = false;
            string currLine = "";

            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];

                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;

                        isBold = false;
                        currLine = "";
                        foreach (DocumentObject docobj in paragraph.ChildObjects)
                        {
                            if (docobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange text = docobj as TextRange;

                                if (text.CharacterFormat.Bold)
                                {
                                    isBold = true;
                                }
                                else
                                {
                                    isBold = false;
                                }
                            }
                        }

                        // contentBox.AppendText(paragraph.ListText);
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];

                            if (cobj.DocumentObjectType == DocumentObjectType.Shape)
                            {
                                ShapeObject shape = cobj as ShapeObject;
                                for (int m = 0; m < shape.ChildObjects.Count; m++)
                                {
                                    if (shape.ChildObjects[m].DocumentObjectType == DocumentObjectType.Paragraph)
                                    {
                                        Paragraph para = shape.ChildObjects[m] as Paragraph;
                                        for (int n = 0; n < para.ChildObjects.Count; n++)
                                        {
                                            if (para.ChildObjects[n].DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange range = para.ChildObjects[n] as TextRange;
                                                string text = range.Text;
                                                isBold = range.CharacterFormat.Bold;
                                                currLine += text;
                                                // contentBox.AppendText(para.ListText + " " + text);

                                                // questionsBox.AppendText(text + Environment.NewLine);

                                            }
                                        }
                                    }
                                }
                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange range = cobj as TextRange;
                                string text = range.Text;
                                //bool isBold = range.CharacterFormat.Bold;
                                currLine += text;
                                // contentBox.AppendText(text);

                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.Table)
                            {
                                Table table = cobj as Table;
                                for (int a = 0; a < table.Rows.Count; a++)
                                {
                                    TableRow row = table.Rows[a];
                                    for (int b = 0; b < row.Cells.Count; b++)
                                    {
                                        TableCell cell = row.Cells[b];
                                        foreach (Paragraph para in cell.Paragraphs)
                                        {
                                            //currLine += para.Text;
                                            //  contentBox.AppendText(para.ListText + " " + para.Text);
                                        }
                                    }
                                }
                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                            {
                                ShapeGroup shapeGroup = cobj as ShapeGroup;
                                for (int k = 0; k < shapeGroup.ChildObjects.Count; k++)
                                {
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Table)
                                    {
                                        Table table = shapeGroup.ChildObjects[k] as Table;
                                        //Table table = cobj as Table;
                                        for (int a = 0; a < table.Rows.Count; a++)
                                        {
                                            TableRow row = table.Rows[a];
                                            for (int b = 0; b < row.Cells.Count; b++)
                                            {
                                                TableCell cell = row.Cells[b];
                                                foreach (Paragraph para in cell.Paragraphs)
                                                {
                                                    //  contentBox.AppendText(para.ListText + " " + para.Text);
                                                }
                                            }
                                        }
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.TextBox)
                                    {
                                        TextBox textbox = shapeGroup.ChildObjects[k] as TextBox;
                                        foreach (DocumentObject objt in textbox.ChildObjects)
                                        {
                                            Console.WriteLine(objt.DocumentObjectType);
                                            //Extract text from paragraph in TextBox.
                                            if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
                                            {
                                                Paragraph para = objt as Paragraph;
                                                //questionsBox.AppendText(para.Text);

                                                isBold = true;

                                                // contentBox.AppendText(para.ListText + para.Text);
                                                currLine += para.Text;
                                            }
                                            if (objt.DocumentObjectType == DocumentObjectType.Table)
                                            {
                                                Table table = objt as Table;
                                                // Table table = cobj as Table;
                                                for (int a = 0; a < table.Rows.Count; a++)
                                                {
                                                    TableRow row = table.Rows[a];
                                                    for (int b = 0; b < row.Cells.Count; b++)
                                                    {
                                                        TableCell cell = row.Cells[b];
                                                        foreach (Paragraph para in cell.Paragraphs)
                                                        {
                                                            //  contentBox.AppendText(para.ListText + para.Text);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        //if(currLine.Trim().Length >1)
                                        //{
                                        //    foreach (KeyValuePair<int, string> entry in headings)
                                        //    {
                                        //        if (currLine.Contains(entry.Value.Trim()))  // .Key must be capitalized
                                        //        {
                                        //            questionsBox.AppendText(currLine + Environment.NewLine);
                                        //        }
                                        //    }
                                        //}

                                    }
                                }
                            }

                        } // outer for loop
                          // contentBox.AppendText(Environment.NewLine);
                        if (isBold)
                        {

                            //headStatus = false;
                            Boolean intStatus = false;
                            foreach (KeyValuePair<int, string> entry in headings)
                            {
                                if (currLine.Contains(entry.Value.Trim()))  // .Key must be capitalized
                                {
                                    intStatus = true;
                                }
                            }
                            if (intStatus == true)
                            {
                                headStatus = true;
                            }
                            else
                            {
                                headStatus = false;
                            }


                            if (currLine.Contains("In questions"))
                            {
                                headStatus = true;
                            }
                            else if (currLine.Contains("EXERCISE"))
                            {
                                headStatus = true;
                            }
                            else if (currLine.Contains("In each of the questions"))
                            {
                                headStatus = true;
                            }
                            else if (currLine.StartsWith("State whether the statements"))
                            {
                                headStatus = true;
                            }
                            // headingsBox.AppendText(currLine + ">>" + headStatus + Environment.NewLine);

                        }

                        if (headStatus == true)
                        {
                            //headingsBox.AppendText(headStatus.ToString() + Environment.NewLine);
                            Paragraph parag = obj as Paragraph;
                            // questionsBox.AppendText(parag.Text + Environment.NewLine);
                            if (!String.IsNullOrEmpty(parag.Text.Trim()))
                            {
                                Paragraph para1 = (Paragraph)parag.Clone();
                                para1.Format.LeftIndent = 30;

                                // para1.Format.ClearFormatting();
                                para1.Format.HorizontalAlignment = SD.HorizontalAlignment.Left;
                                s2.Paragraphs.Add(para1);
                            }

                        }

                    } //if paragraph

                }
            }



            doc2.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.docx", FileFormat.Docx);

            //doc2.SaveToFile(@"C: \Users\C K Bhushan\Desktop\NCERT Books\Exemplar- phase-2.docx", FileFormat.Docx);
            doc2.Close();
            doc.Close();
        }


        public void imageTags()
        {
            var sImgName = Path.GetFileNameWithoutExtension(LoadFilePath);
            string imagePath = sImgName + "-images/";
            Document doc = new Document();
            doc.LoadFromFile(LoadFilePath);
            int i = 1;
            foreach (Section sec in doc.Sections)
            {
                foreach (Paragraph para in sec.Paragraphs)
                {
                    List<DocumentObject> pictures = new List<DocumentObject>();
                    List<DocumentObject> oleObjects = new List<DocumentObject>();
                    foreach (DocumentObject dobjt in para.ChildObjects)
                    {
                        if (dobjt.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            pictures.Add(dobjt);
                        }
                        //if(dobjt.DocumentObjectType == DocumentObjectType.OleObject)
                        // {
                        //     oleObjects.Add(dobjt);
                        // }
                    }
                    foreach (DocumentObject pic in pictures)
                    {
                        int index = para.ChildObjects.IndexOf(pic);
                        TextRange range = new TextRange(doc);
                        //string imgTextReplace = @" < img src=""images/image001.jpg"" alt=""images""/>";
                        range.Text = string.Format(@"<img>https://repository.evelynlearning.com/imageresource/" + imagePath + sImgName + "-image00" + i.ToString() + @".png</img>");
                        para.ChildObjects.Insert(index, range);
                        para.ChildObjects.Remove(pic);
                        i++;
                    }
                }
            }

            doc.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-1.docx", FileFormat.Docx);

        }

        public void extractImages()
        {
            int fileCount = 0;
            string imageName;
            //string imageDir = subjectBox.SelectedItem.ToString();


            //int i = 0;
            try
            {
                fileCount = fileCount + 1;

                Document doc = new Document();
                doc.LoadFromFile(LoadFilePath);
                List<DocPicture> DocPictureList = new List<DocPicture>();
                List<DocPicture> mathPictureList = new List<DocPicture>();
                List<DocPicture> image = new List<DocPicture>();
                image.Clear();
                DocPictureList.Clear();
                mathPictureList.Clear();
                var sImgName = Path.GetFileNameWithoutExtension(LoadFilePath);
                string imagePath = Path.GetDirectoryName(LoadFilePath) + "/" + sImgName + "-images/";
                if (!Directory.Exists(imagePath))
                {
                    Directory.CreateDirectory(imagePath);
                }

                //Loop through contents
                foreach (Section section in doc.Sections)
                {
                    foreach (DocumentObject obj in section.Body.ChildObjects)
                    {
                        if (obj is Paragraph)
                        {
                            Paragraph para = obj as Paragraph;
                            foreach (DocumentObject cobj in para.ChildObjects)
                            {
                                //Find DocPicture object and add it in DocPictureList
                                if (cobj is DocPicture)
                                {
                                    DocPicture pic = cobj as DocPicture;
                                    DocPictureList.Add(pic);

                                }
                                //Find DocOleObject object and add it in mathPictureList
                                if (cobj is DocOleObject)
                                {
                                    DocOleObject ole = cobj as DocOleObject;
                                    mathPictureList.Add(ole.OlePicture);
                                }
                            }
                            image = DocPictureList.Except(mathPictureList).ToList();
                            //image = mathPictureList;

                        }
                    }

                    int imageIdx = 1;
                    foreach (DocPicture pic in DocPictureList)
                    {
                        // textBox1.AppendText(pic + Environment.NewLine);
                        //imageName = string.Format(sImgName + "_" + DateTime.Now.ToString("HH_mm_ss") + "_image00{0}.png", imageIdx);

                        imageName = string.Format(sImgName + "_image00{0}.png", imageIdx);
                        pic.Image.Save(imagePath + imageName, System.Drawing.Imaging.ImageFormat.Png);
                        imageIdx += 1;

                    }

                }

                image.Clear();
                DocPictureList.Clear();
                mathPictureList.Clear();
                doc.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // IssueBox1.AppendText("File is not in expected format" + " " + ex + " " + Environment.NewLine);
            }
            // MessageBox.Show("Done Files");

        }

        public void disp()
        {
            foreach (KeyValuePair<int, string> entry in headings)
            {
                Console.WriteLine(entry.Key + "-" + entry.Value);
            }
        }

        public void initComp()
        {
            int counter = 0;
            string[] array = parameters.Split('&');
            foreach (string value in array)
            {
                if (value.Contains(".docx") || value.Contains(".Docx") || value.Contains(".DOCX"))
                {
                    if (value.Contains("Answers") || value.Contains("answers"))
                    {
                        AnswerFilePath = value;
                    }
                    else
                    {
                        LoadFilePath = value;
                    }
                }
                else
                {
                    headings.Add(counter, value);
                    counter += 1;
                }

            }

            string FILE_NAME = "headings.txt";
            string currLine = "";


            if (System.IO.File.Exists(FILE_NAME) == true)
            {
                System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);

                while (objReader.Peek() != -1)
                {
                    currLine = objReader.ReadLine();
                    headings.Add(counter, currLine);
                    counter += 1;
                }
                objReader.Dispose();
            }

            HeadingsCounter = counter;
        }

        public void CreateDoc()
        {
            //Create a document object
            Document doc = new Document();

            //Add a section
            Section section = doc.AddSection();

            //Add a paragrah
            Paragraph paragraph = section.AddParagraph();

            //Append text to the paragraph
            paragraph.AppendText("This article shows you how to mannually add Spire.Doc as dependency in a .NET Core application.");

            //Save to file
            doc.SaveToFile(@"D:\Output.docx", FileFormat.Docx2013);
            Console.WriteLine("Word creatino done.");
            Console.ReadLine();
            doc.Close();
        }



    }
}
