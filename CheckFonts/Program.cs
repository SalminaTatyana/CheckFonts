using System;
using System.IO;
using System.Text.RegularExpressions;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CheckFonts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string fileNameResult = "";

            try
            {
                if (args.Length > 0)
                {
                    if (args.Length != 2)
                    {
                        Console.WriteLine("Неверное количество файлов. Проверьте наличие строго 2 файлов в bat файле");
                        Console.ReadKey();
                    }
                    else
                    {
                        Document doc = OpenDoc(args[0]);
                        string[] name = args[1].Split('.');

                        string expansion = $".{name[name.Length - 1]}";
                        if (expansion == ".txt")
                        {
                            fileNameResult = args[1];
                        }
                        else
                        {
                            fileNameResult = $"{args[1]}.txt";
                        }

                        WriteFile($"Проверяемый файл {args[0]}, выходной файл {args[1]}", fileNameResult);

                        string wrongFonts = WrongFonts(doc);
                        if (!string.IsNullOrEmpty(wrongFonts))
                        {
                            WriteFile(wrongFonts, fileNameResult);
                            Console.WriteLine(wrongFonts);

                        }
                        else
                        {
                            WriteFile("Проверка успешно завершена", fileNameResult);
                            Console.WriteLine("Проверка успешно завершена");
                        }

                        Console.ReadKey();

                    }
                }
                else
                {
                    Console.WriteLine("Не заданы файлы. Проверьте наличие строго 2 файлов в bat файле");
                    Console.ReadKey();
                    //Document doc = OpenDoc("1.docx");
                    //string[] name = "1".Split('.');

                    //string expansion = $".{name[name.Length - 1]}";
                    //if (expansion == ".txt")
                    //{
                    //    fileNameResult = "1";
                    //}
                    //else
                    //{
                    //    fileNameResult = $"1.txt";
                    //}

                    //WriteFile($"Проверяемый файл 1.docx, выходной файл 1.txt", fileNameResult);

                    //string wrongFonts = WrongFonts(doc);
                    //if (!string.IsNullOrEmpty(wrongFonts))
                    //{
                    //    WriteFile(wrongFonts, fileNameResult);
                    //    Console.WriteLine(wrongFonts);

                    //}
                    //else
                    //{
                    //    WriteFile("Проверка успешно завершена", fileNameResult);
                    //    Console.WriteLine("Проверка успешно завершена");
                    //}

                    //Console.ReadKey();
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Что-то пошло не так!");
                Console.ReadKey();
            }


            Document OpenDoc(string path)
            {
                if (File.Exists(path))
                {
                    Document document = new Document(path.Trim());
                    return document;
                }
                else
                {

                    WriteFile("Файла не существует", fileNameResult);
                    Console.WriteLine("Файла не существует");
                    Console.ReadKey();
                    return null;
                }
            }

            string WrongFonts(Document doc)
            {
                bool isCheck = false;
                string isWrong = "";
                foreach (Section sec in doc.Sections)
                {

                    foreach (DocumentObject obj in sec.Body.ChildObjects)
                    {
                        if (obj is Paragraph)
                        {
                            var para = obj as Paragraph;

                            foreach (DocumentObject Pobj in para.ChildObjects)
                            {
                                if (Pobj is TextRange)
                                {
                                    TextRange textRange = Pobj as TextRange;
                                    if (Regex.IsMatch(textRange.Text.Trim(), "^введение$",
                                            RegexOptions.IgnoreCase))
                                    {
                                        isCheck = true;
                                    }
                                    
                                    if (isCheck)
                                    {

                                        //Обычные абзацы
                                        if (!Regex.IsMatch(textRange.Text.Trim(),
                                                "Рисунок [1-9]\\.[0-9]{0,9}|рисунок [1-9]\\.[0-9]{0,9}|string |int |bool |for |if |while |foreach |{|}|\\);|Дата обращения|Select|join|where|insert|update|order|and|end|or|<|>|from",
                                                RegexOptions.IgnoreCase))
                                        {
                                            if ((textRange.CharacterFormat.FontName != "Times New Roman" ||
                                                 textRange.CharacterFormat.FontSize != 14 ||
                                                 para.Format.FirstLineIndent <= 0) &&
                                                !string.IsNullOrEmpty(textRange.Text)&&textRange.CharacterFormat.LocaleIdASCII!= 1033)
                                            {
                                                if (textRange.CharacterFormat.FontName != "Times New Roman")
                                                {
                                                    if (!isWrong.Contains("Неверный шрифт в абзаце: " + para.Text + "\n"))
                                                    {
                                                        isWrong += "Неверный шрифт в абзаце: " + para.Text + "\n";
                                                    }
                                                   
                                                }

                                                if (textRange.CharacterFormat.FontSize != 14)
                                                {
                                                    if (!isWrong.Contains("Неверный размер шрифта в абзаце: " + para.Text))
                                                    {
                                                        isWrong += "Неверный размер шрифта в абзаце: " + para.Text +
                                                                   "\n";
                                                    }
                                                   
                                                }

                                                if (para.Format.FirstLineIndent <= 0)
                                                {
                                                    if (!isWrong.Contains("Неверный абзацный отступ в абзаце: " + para.Text))
                                                    {
                                                        isWrong += "Неверный абзацный отступ в абзаце: " + para.Text +
                                                                   "\n";
                                                    }
                                                   

                                                }

                                            }
                                        }

                                        //Абзацы с рисунком
                                        if (Regex.IsMatch(textRange.Text.Trim(), "Рисунок [1-9]\\.[0-9]{0,9}",
                                                RegexOptions.IgnoreCase) && !Regex.IsMatch(textRange.Text,
                                                "Дата обращения",
                                                RegexOptions.IgnoreCase) &&
                                            !string.IsNullOrEmpty(textRange.Text))
                                        {
                                            if (textRange.CharacterFormat.FontName != "Times New Roman" ||
                                                textRange.CharacterFormat.FontSize != 14 ||
                                                para.Format.FirstLineIndent != 0)
                                            {
                                                if (textRange.CharacterFormat.FontName != "Times New Roman")
                                                {
                                                    if (!isWrong.Contains("Неверный шрифт в абзаце с рисунком: " + para.Text))
                                                    {
                                                        isWrong += "Неверный шрифт в абзаце с рисунком: " + para.Text +
                                                                   "\n";
                                                    }
                                                    
                                                }

                                                if (textRange.CharacterFormat.FontSize != 14)
                                                {
                                                    if (!isWrong.Contains("Неверный размер шрифта в абзаце с рисунком: " +
                                                                          para.Text))
                                                    {
                                                        isWrong += "Неверный размер шрифта в абзаце с рисунком: " +
                                                                   para.Text + "\n";
                                                    }
                                                   
                                                }

                                                if (para.Format.FirstLineIndent <= 0)
                                                {
                                                    if (!isWrong.Contains("Неверный абзацный отступ в абзаце с рисунком: " +
                                                                          para.Text))
                                                    {
                                                        isWrong += "Неверный абзацный отступ в абзаце с рисунком: " +
                                                                   para.Text + "\n";
                                                    }
                                                    
                                                }

                                            }
                                        }

                                        //Абзацы с кодом
                                        if (Regex.IsMatch(textRange.Text,
                                                "string |int |bool |for |if |while |foreach |{|}|^\\);|Select|join|where|insert|update|order|and|end|or|<|>|from",
                                                RegexOptions.IgnoreCase) && !Regex.IsMatch(textRange.Text,
                                                "Дата обращения",
                                                RegexOptions.IgnoreCase) &&
                                            !string.IsNullOrEmpty(textRange.Text))
                                        {
                                            if (
                                                textRange.CharacterFormat.FontSize != 10 ||
                                                para.Format.FirstLineIndent != 0)
                                            {


                                                if (textRange.CharacterFormat.FontSize != 10)
                                                {
                                                    if (!isWrong.Contains(
                                                            "Неверный размер шрифта в абзаце с кодом: " +
                                                            para.Text))
                                                    {
                                                        isWrong += "Неверный размер шрифта в абзаце с кодом: " +
                                                                   para.Text + "\n";
                                                    }
                                                       
                                                }

                                                if (para.Format.FirstLineIndent <= 0)
                                                {
                                                    if (!isWrong.Contains(
                                                            "Неверный абзацный отступ в абзаце с кодом: " +
                                                            para.Text))
                                                    {
                                                        isWrong += "Неверный абзацный отступ в абзаце с кодом: " +
                                                                   para.Text + "\n";
                                                    }
                                                    
                                                }

                                            }
                                        }

                                    }
                                }
                            }

                        }
                    }

                }

                return isWrong;
            }
        }

        static public void WriteFile(string text, string fileNameResult)
        {
            using (StreamWriter writer = new StreamWriter(fileNameResult, true, System.Text.Encoding.Default))
            {
                writer.WriteLine(text);
            }

        }
    }
}
