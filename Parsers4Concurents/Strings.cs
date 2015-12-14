using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parsers4Сompetitor
{
    static class Strings
    {
        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее заданное количество символов.
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает все подстроки после найденных меток
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="col_ch"></param>
        /// <returns></returns>
        public static List<string> GetWithIn(string str1, string str2, int col_ch)
        {
            List<string> rez = new List<string>();
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    if (col_ch > 0)
                    {
                        string str = str1.Substring(i + str2.Length, col_ch);
                        rez.Add(str);
                    }
                    else  // обработка отрицательного значения col_ch при отмотке
                    {
                        i += col_ch;
                        int col = col_ch - col_ch * 2;
                        string str = str1.Substring(i, col);
                        rez.Add(str);
                        i += col;
                    }

                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее символы до заданного (ch).
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает все подстроки после найденных меток
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="ch"></param>
        /// <returns></returns>
        public static List<string> GetWithInCh(string str1, string str2, char ch)
        {
            List<string> rez = new List<string>();
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    //выбор до опред. симвовла
                    //new string(s.TakeWhile(x => x != '{').ToArray()).Length
                    string str11 = str1.Substring(i + str2.Length, str1.Length - (i + str2.Length));
                    string str = new string(str11.TakeWhile(x_i => x_i != ch).ToArray());
                    rez.Add(str);
                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее символы до заданного (ch).
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает первую строку
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="ch"></param>
        /// <returns></returns>
        public static string GetWithInChString(string str1, string str2, char ch)
        {
            string rez = "";
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    //выбор до опред. симвовла
                    //new string(s.TakeWhile(x => x != '{').ToArray()).Length
                    string str11 = str1.Substring(i + str2.Length, str1.Length - (i + str2.Length));
                    rez = new string(str11.TakeWhile(x_i => x_i != ch).ToArray());
                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее символы до str_end.
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает все подстроки после найденных меток
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="ch"></param>
        /// <returns></returns>
        public static List<string> GetWithInStrList(string str1, string str2, string str_end)
        {
            List<string> rez = new List<string>();
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    int i_str = str1.IndexOf(str_end, i + str2.Length);
                    if (i_str != -1)
                    {
                        int l_str1 = str1.Length;
                        int l_str2 = str2.Length;
                        int str_search = i_str - (i + l_str2);
                        string str = str1.Substring(i + l_str2, str_search);
                        rez.Add(str);
                    }
                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит первое вхождение метки str2, и получает после нее символы до метки str_end.
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="str_end"></param>
        /// <returns></returns>
        public static string GetWithInStr(string str1, string str2, string str_end)
        {
            string rez = "";
            int i = 0;
            i = str1.IndexOf(str2, 0); // получаем индекс первого вхождения с 0-го индекса
            if (i > -1)
            {
                int i_str2 = str1.IndexOf(str_end, i + str2.Length);
                if (i_str2 != -1)
                {
                    int l_str1 = str1.Length;
                    int l_str2 = str2.Length;
                    int str_search = i_str2 - (i + l_str2);
                    rez = str1.Substring(i + l_str2, str_search);
                }
            }
            return rez;
        }
    }
}
