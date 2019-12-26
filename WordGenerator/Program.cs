using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace WordGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start ...");

            //EXCEL
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            Workbook excelDoc = excelApp.Workbooks.Open("D:\\A-DATA\\GoogleDriveYork\\PMCS\\[001] Essay\\Scores.xlsx");

            //WORD
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;

            try
            {
                // EXCEL
                Worksheet mainSheet = (Worksheet)excelDoc.Worksheets.Item[1];
                Worksheet q1Sheet = (Worksheet)excelDoc.Worksheets.Item[2];
                Worksheet q3Sheet = (Worksheet)excelDoc.Worksheets.Item[3];

                int studentIdCol = 2;
                int q1Threshold = 71;
                int q1AnswerCol = 13;
                int q1FeedbackCol = 14;
                int q3AnswerCol = 14;
                int q3FeedbackCol = 15;

                int rowFrom = 2;
                int rowTo = 139;
                for (int row = rowFrom; row <= rowTo; row++)
                {
                    Microsoft.Office.Interop.Excel.Range studentIdRange = (Microsoft.Office.Interop.Excel.Range)mainSheet.Cells[row, studentIdCol];
                    String studentId = (String)studentIdRange.Text;
                    Console.WriteLine(studentId);
                    if (studentId.Equals(""))
                        continue;

                    Microsoft.Office.Interop.Excel.Range q1AnswerRange = (Microsoft.Office.Interop.Excel.Range)q1Sheet.Cells[row, q1AnswerCol];
                    String q1Answer = (String)q1AnswerRange.Text;

                    Microsoft.Office.Interop.Excel.Range q1FeedbackRange = (Microsoft.Office.Interop.Excel.Range)q1Sheet.Cells[row, q1FeedbackCol];
                    String q1Feedback = (String)q1FeedbackRange.Text;

                    Microsoft.Office.Interop.Excel.Range q3AnswerRange = (Microsoft.Office.Interop.Excel.Range)q3Sheet.Cells[row, q3AnswerCol];
                    String q3Answer = (String)q3AnswerRange.Text;

                    Microsoft.Office.Interop.Excel.Range q3FeedbackRange = (Microsoft.Office.Interop.Excel.Range)q3Sheet.Cells[row, q3FeedbackCol];
                    String q3Feedback = (String)q3FeedbackRange.Text;


                    String path = "D:\\A-DATA\\GoogleDriveYork\\PMCS\\[001] Essay\\scores\\" + studentId + ".docx";
                    Document wordDoc = wordApp.Documents.Open(path);

                    try
                    {
                        //for (int i = 1; i < document.Paragraphs.Count; i++)
                        //{
                        //    String text = document.Paragraphs[i].Range.Text;
                        //    Console.WriteLine(i + ": " + text);
                        //}

                        String temp = wordDoc.Paragraphs[4].Range.Text;
                        if (!q1Answer.Equals("0"))
                        {
                            temp = temp.Replace("Q1", q1Answer);
                        }

                        //String q2 = temp.Substring(18, 2);
                        if (!q3Answer.Equals("0"))
                        {
                            temp = temp.Replace("Q3", q3Answer);
                        }

                        wordDoc.Paragraphs[4].Range.Text = temp;

                        if (!q1Answer.Equals("0"))
                        {
                            //0 – 12  13 – 16 17 – 24 25 – 32 33 – 40
                            if (Double.Parse(q1Answer) >= 33)
                            {
                                wordDoc.Paragraphs[29].Range.Text = q1Answer;
                            }
                            else if (Double.Parse(q1Answer) >= 25)
                            {
                                wordDoc.Paragraphs[28].Range.Text = q1Answer;
                            }
                            else if (Double.Parse(q1Answer) >= 17)
                            {
                                wordDoc.Paragraphs[27].Range.Text = q1Answer;
                            }
                            else if (Double.Parse(q1Answer) >= 13)
                            {
                                wordDoc.Paragraphs[26].Range.Text = q1Answer;
                            }
                            else if (Double.Parse(q1Answer) > 0)
                            {
                                wordDoc.Paragraphs[25].Range.Text = q1Answer;
                            }
                            wordDoc.Paragraphs[32].Range.Text = q1Feedback;
                        }

                        if (!q3Answer.Equals("0"))
                        {
                            //0 – 19  10 – 12 13 – 18 19 – 24 25 – 30
                            if (q1Answer.Equals("0"))
                            {
                                if (Double.Parse(q3Answer) >= 25)
                                {
                                    wordDoc.Paragraphs[90].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) >= 19)
                                {
                                    wordDoc.Paragraphs[89].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) >= 13)
                                {
                                    wordDoc.Paragraphs[88].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) >= 10)
                                {
                                    wordDoc.Paragraphs[87].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) > 0)
                                {
                                    wordDoc.Paragraphs[86].Range.Text = q3Answer;
                                }
                                wordDoc.Paragraphs[93].Range.Text = q3Feedback;
                            }

                            if (!q1Answer.Equals("0"))
                            {
                                if (Double.Parse(q3Answer) >= 25)
                                {
                                    wordDoc.Paragraphs[89].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) >= 19)
                                {
                                    wordDoc.Paragraphs[88].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) >= 13)
                                {
                                    wordDoc.Paragraphs[87].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) >= 10)
                                {
                                    wordDoc.Paragraphs[86].Range.Text = q3Answer;
                                }
                                else if (Double.Parse(q3Answer) > 0)
                                {
                                    wordDoc.Paragraphs[85].Range.Text = q3Answer;
                                }
                                wordDoc.Paragraphs[92].Range.Text = q3Feedback;
                            }
                        }

                        wordDoc.Save();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {
                        wordDoc.Close();

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                wordApp.Quit();
                excelDoc.Close();
                excelApp.Quit();
            }

            Console.WriteLine("Finished!");
        }
    }
}
