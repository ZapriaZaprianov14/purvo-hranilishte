using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;

namespace DFN2
{
    public class Helper
    {
        public static string CnnVaL(string name)
        {
            return ConfigurationManager.ConnectionStrings[name].ConnectionString;
        }
        public static void AddRange(int from, int to, List<int> list)
        {
            if (list.Count == 0)
            {
                int numberOfElements = to - from + 1;
                for (int i = 0; i < numberOfElements; i++)
                {
                    list.Add(from + i);
                }
            }
            else
            {
                list.Add(0);
                int numberOfElements = to - from + 1;
                for (int i = 0; i < numberOfElements; i++)
                {
                    list.Add(from + i);
                }
            }
        }
        public static void ResetRange(int from, int to, List<int> list)
        {
            bool filled = false;
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] == from - 1)
                {
                    for (int j = 0; j < to - from + 1; j++)
                    {
                        list[i + j + 1] = from + j;
                    }
                    filled = true;
                    break;
                }
            }
            if (!filled)
            {
                for (int j = 0; j < to - from + 1; j++)
                {
                    list[j] = from + j;
                }
            }
        }
        public static void RemoveRange(int from, int to, List<int> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] == from)
                {
                    for (int j = i; true; j++)
                    {
                        if (list[j] != to)
                        {
                            list[j] = 0;
                        }
                        else
                        {
                            list[j] = 0;
                            break;
                        }
                    }
                    break;
                }
            }
        }
        public static void CutList(List<int> list, List<List<int>> listOfRanges)
        {
            List<int> helpList = new List<int>();
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] != 0)
                {
                    helpList.Add(list[i]);
                }
                else
                {
                    if (helpList.Count > 0)
                    {
                        List<int> second = new List<int>();
                        for (int j = 0; j < helpList.Count; j++)
                        {
                            second.Add(helpList[j]);
                        }
                        listOfRanges.Add(second);
                        helpList.Clear();
                    }
                }
            }
            if (helpList.Count > 0)
            {
                List<int> second = new List<int>();
                for (int j = 0; j < helpList.Count; j++)
                {
                    second.Add(helpList[j]);
                }
                listOfRanges.Add(second);
                helpList.Clear();
            }

        }
        public static bool CheckIfRangeIsCorrect(string from, string to)
        {
            if (Int32.TryParse(from, out int result) == false || Int32.TryParse(to, out result) == false || int.Parse(from) > int.Parse(to))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public static string MakeTo6(string tb)
        {
            int n1 = 0;
            string result = "";
            bool correct = int.TryParse(tb, out n1);
            if (correct)
                result = n1.ToString().PadLeft(6, '0');
            return result;
        }
        public static List<string> CutStringList(List<string> list)
        {
            List<string> helpList = new List<string>();
            List<string> result = new List<string>();
            List<List<string>> listOfRanges = new List<List<string>>();
            if (list.Count == 1)
            {
                result.Add(list[0] + "-" + list[0]);
                return result;
            }
            else if (list.Count == 0)
            {
                return result;
            }
            else
            {
                for (int i = 0; i < list.Count - 1; i++)
                {
                    if (int.Parse(list[i]) + 1 == int.Parse(list[i + 1]))
                    {
                        helpList.Add(list[i]);
                    }
                    else
                    {
                        helpList.Add(list[i]);

                        List<string> second = new List<string>();
                        for (int j = 0; j < helpList.Count; j++)
                        {
                            second.Add(helpList[j]);
                        }
                        listOfRanges.Add(second);
                        helpList.Clear();

                    }
                }
            }
            if (helpList.Count > 0)
            {
                List<string> second = new List<string>();
                for (int j = 0; j < helpList.Count; j++)
                {
                    second.Add(helpList[j]);
                    second.Add(list[list.Count - 1]);
                }
                listOfRanges.Add(second);
                helpList.Clear();
            }
            if (int.Parse(list[list.Count - 1]) - 1 != int.Parse(list[list.Count - 2]))
            {
                helpList.Add(list[list.Count - 1]);
                listOfRanges.Add(helpList);
            }
            foreach (var element in listOfRanges)
            {
                result.Add(element[0] + "-" + element[element.Count - 1]);
            }
            return result;
        }
    }
}
public static class Globals
{
    public static string Tab1RegBookDoc1;
    public static string Tab1CanceledDoc1;
    public static string Tab1ForDesDoc1;
    public static string Tab1ReceivedDoc1;
    public static string Tab1Doc1Destroyed;
    public static string Tab1Doc1Rest;
    public static string Tab1Doc1Start;
    public static string Tab1Doc1End;
    public static string Doc1CounterTBox;

    public static SqlConnection connection;
    public static string OrderNumber;
    public static DateTime OrderDate;

    public static string Tab1RegBookDoc2;
    public static string Tab1CanceledDoc2;
    public static string Tab1ForDesDoc2;
    public static string Tab1ReceivedDoc2;
    public static string Tab1Doc2Destroyed;
    public static string Tab1Doc2Rest;
    public static string Tab1Doc2Start;
    public static string Tab1Doc2End;
    public static string Doc2CounterTBox;

    public static string Tab1RegBookDoc3;
    public static string Tab1CanceledDoc3;
    public static string Tab1ForDesDoc3;
    public static string Tab1ReceivedDoc3;
    public static string Tab1Doc3Destroyed;
    public static string Tab1Doc3Rest;
    public static string Tab1Doc3Start;
    public static string Tab1Doc3End;
    public static string Doc3CounterTBox;

    public static string Tab1RegBookDoc4;
    public static string Tab1CanceledDoc4;
    public static string Tab1ForDesDoc4;
    public static string Tab1ReceivedDoc4;
    public static string Tab1Doc4Destroyed;
    public static string Tab1Doc4Rest;
    public static string Tab1Doc4Start;
    public static string Tab1Doc4End;
    public static string Doc4CounterTBox;

    public static string Tab1RegBookDoc5;
    public static string Tab1CanceledDoc5;
    public static string Tab1ForDesDoc5;
    public static string Tab1ReceivedDoc5;
    public static string Tab1Doc5Destroyed;
    public static string Tab1Doc5Rest;
    public static string Tab1Doc5Start;
    public static string Tab1Doc5End;
    public static string Doc5CounterTBox;

    public static string Tab1RegBookDoc6;
    public static string Tab1CanceledDoc6;
    public static string Tab1ForDesDoc6;
    public static string Tab1ReceivedDoc6;
    public static string Tab1Doc6Destroyed;
    public static string Tab1Doc6Rest;
    public static string Tab1Doc6Start;
    public static string Tab1Doc6End;
    public static string Doc6CounterTBox;

    public static string Tab1RegBookDoc7;
    public static string Tab1CanceledDoc7;
    public static string Tab1ForDesDoc7;
    public static string Tab1ReceivedDoc7;
    public static string Tab1Doc7Destroyed;
    public static string Tab1Doc7Rest;
    public static string Tab1Doc7Start;
    public static string Tab1Doc7End;
    public static string Doc7CounterTBox;

    public static string Tab1RegBookDoc8;
    public static string Tab1CanceledDoc8;
    public static string Tab1ForDesDoc8;
    public static string Tab1ReceivedDoc8;
    public static string Tab1Doc8Destroyed;
    public static string Tab1Doc8Rest;
    public static string Tab1Doc8Start;
    public static string Tab1Doc8End;
    public static string Doc8CounterTBox;

    public static string Tab1RegBookDoc9;
    public static string Tab1CanceledDoc9;
    public static string Tab1ForDesDoc9;
    public static string Tab1ReceivedDoc9;
    public static string Tab1Doc9Destroyed;
    public static string Tab1Doc9Rest;
    public static string Tab1Doc9Start;
    public static string Tab1Doc9End;
    public static string Doc9CounterTBox;

    public static string Tab1RegBookDoc10;
    public static string Tab1CanceledDoc10;
    public static string Tab1ForDesDoc10;
    public static string Tab1ReceivedDoc10;
    public static string Tab1Doc10Destroyed;
    public static string Tab1Doc10Rest;
    public static string Tab1Doc10Start;
    public static string Tab1Doc10End;
    public static string Doc10CounterTBox;

    public static string Tab1RegBookDoc11;
    public static string Tab1CanceledDoc11;
    public static string Tab1ForDesDoc11;
    public static string Tab1ReceivedDoc11;
    public static string Tab1Doc11Destroyed;
    public static string Tab1Doc11Rest;
    public static string Tab1Doc11Start;
    public static string Tab1Doc11End;
    public static string Doc11CounterTBox;

    public static string Tab1RegBookDoc12;
    public static string Tab1CanceledDoc12;
    public static string Tab1ForDesDoc12;
    public static string Tab1ReceivedDoc12;
    public static string Tab1Doc12Destroyed;
    public static string Tab1Doc12Rest;
    public static string Tab1Doc12Start;
    public static string Tab1Doc12End;
    public static string Doc12CounterTBox;

    public static string Tab1RegBookDoc13;
    public static string Tab1CanceledDoc13;
    public static string Tab1ForDesDoc13;
    public static string Tab1ReceivedDoc13;
    public static string Tab1Doc13Destroyed;
    public static string Tab1Doc13Rest;
    public static string Tab1Doc13Start;
    public static string Tab1Doc13End;
    public static string Doc13CounterTBox;

    public static string Tab1RegBookDoc14;
    public static string Tab1CanceledDoc14;
    public static string Tab1ForDesDoc14;
    public static string Tab1ReceivedDoc14;
    public static string Tab1Doc14Destroyed;
    public static string Tab1Doc14Rest;
    public static string Tab1Doc14Start;
    public static string Tab1Doc14End;
    public static string Doc14CounterTBox;

    public static string Tab2RegBookDoc1;
    public static string Tab2CanceledDoc1;
    public static string Tab2ForDesDoc1;
    public static string Tab2ReceivedDoc1;
    public static string Tab2Doc1Destroyed;
    public static string Tab2Doc1Rest;
    public static string Tab2Doc1Start;
    public static string Tab2Doc1End;

    public static string Tab2RegBookDoc2;
    public static string Tab2CanceledDoc2;
    public static string Tab2ForDesDoc2;
    public static string Tab2ReceivedDoc2;
    public static string Tab2Doc2Destroyed;
    public static string Tab2Doc2Rest;
    public static string Tab2Doc2Start;
    public static string Tab2Doc2End;

    public static string Tab2RegBookDoc3;
    public static string Tab2CanceledDoc3;
    public static string Tab2ForDesDoc3;
    public static string Tab2ReceivedDoc3;
    public static string Tab2Doc3Destroyed;
    public static string Tab2Doc3Rest;
    public static string Tab2Doc3Start;
    public static string Tab2Doc3End;


    public static string Tab2RegBookDoc4;
    public static string Tab2CanceledDoc4;
    public static string Tab2ForDesDoc4;
    public static string Tab2ReceivedDoc4;
    public static string Tab2Doc4Destroyed;
    public static string Tab2Doc4Rest;
    public static string Tab2Doc4Start;
    public static string Tab2Doc4End;


    public static string Tab2RegBookDoc5;
    public static string Tab2CanceledDoc5;
    public static string Tab2ForDesDoc5;
    public static string Tab2ReceivedDoc5;
    public static string Tab2Doc5Destroyed;
    public static string Tab2Doc5Rest;
    public static string Tab2Doc5Start;
    public static string Tab2Doc5End;

    public static string Tab2RegBookDoc6;
    public static string Tab2CanceledDoc6;
    public static string Tab2ForDesDoc6;
    public static string Tab2ReceivedDoc6;
    public static string Tab2Doc6Destroyed;
    public static string Tab2Doc6Rest;
    public static string Tab2Doc6Start;
    public static string Tab2Doc6End;

    public static string Tab2RegBookDoc7;
    public static string Tab2CanceledDoc7;
    public static string Tab2ForDesDoc7;
    public static string Tab2ReceivedDoc7;
    public static string Tab2Doc7Destroyed;
    public static string Tab2Doc7Rest;
    public static string Tab2Doc7Start;
    public static string Tab2Doc7End;

    public static string Tab2RegBookDoc8;
    public static string Tab2CanceledDoc8;
    public static string Tab2ForDesDoc8;
    public static string Tab2ReceivedDoc8;
    public static string Tab2Doc8Destroyed;
    public static string Tab2Doc8Rest;
    public static string Tab2Doc8Start;
    public static string Tab2Doc8End;

    public static string Tab2RegBookDoc9;
    public static string Tab2CanceledDoc9;
    public static string Tab2ForDesDoc9;
    public static string Tab2ReceivedDoc9;
    public static string Tab2Doc9Destroyed;
    public static string Tab2Doc9Rest;
    public static string Tab2Doc9Start;
    public static string Tab2Doc9End;

    public static string Tab2RegBookDoc10;
    public static string Tab2CanceledDoc10;
    public static string Tab2ForDesDoc10;
    public static string Tab2ReceivedDoc10;
    public static string Tab2Doc10Destroyed;
    public static string Tab2Doc10Rest;
    public static string Tab2Doc10Start;
    public static string Tab2Doc10End;

    public static string Tab2RegBookDoc11;
    public static string Tab2CanceledDoc11;
    public static string Tab2ForDesDoc11;
    public static string Tab2ReceivedDoc11;
    public static string Tab2Doc11Destroyed;
    public static string Tab2Doc11Rest;
    public static string Tab2Doc11Start;
    public static string Tab2Doc11End;

    public static string Tab2RegBookDoc12;
    public static string Tab2CanceledDoc12;
    public static string Tab2ForDesDoc12;
    public static string Tab2ReceivedDoc12;
    public static string Tab2Doc12Destroyed;
    public static string Tab2Doc12Rest;
    public static string Tab2Doc12Start;
    public static string Tab2Doc12End;

    public static string Tab2RegBookDoc13;
    public static string Tab2CanceledDoc13;
    public static string Tab2ForDesDoc13;
    public static string Tab2ReceivedDoc13;
    public static string Tab2Doc13Destroyed;
    public static string Tab2Doc13Rest;
    public static string Tab2Doc13Start;
    public static string Tab2Doc13End;

    public static string Tab2RegBookDoc14;
    public static string Tab2CanceledDoc14;
    public static string Tab2ForDesDoc14;
    public static string Tab2ReceivedDoc14;
    public static string Tab2Doc14Destroyed;
    public static string Tab2Doc14Rest;
    public static string Tab2Doc14Start;
    public static string Tab2Doc14End;

    public static string Tab3Doc1Start;
    public static string Tab3Doc1End;
    public static string Tab3Doc1Count;

    public static string Tab3Doc2Start;
    public static string Tab3Doc2End;
    public static string Tab3Doc2Count;

    public static string Tab3Doc3Start;
    public static string Tab3Doc3End;
    public static string Tab3Doc3Count;

    public static string Tab3Doc4Start;
    public static string Tab3Doc4End;
    public static string Tab3Doc4Count;

    public static string Tab3Doc5Start;
    public static string Tab3Doc5End;
    public static string Tab3Doc5Count;

    public static string Tab3Doc6Start;
    public static string Tab3Doc6End;
    public static string Tab3Doc6Count;

    public static string Tab3Doc7Start;
    public static string Tab3Doc7End;
    public static string Tab3Doc7Count;

    public static string Tab3Doc8Start;
    public static string Tab3Doc8End;
    public static string Tab3Doc8Count;

    public static string Tab3Doc9Start;
    public static string Tab3Doc9End;
    public static string Tab3Doc9Count;

    public static string Tab3Doc10Start;
    public static string Tab3Doc10End;
    public static string Tab3Doc10Count;

    public static string Tab3Doc11Start;
    public static string Tab3Doc11End;
    public static string Tab3Doc11Count;

    public static string Tab3Doc12Start;
    public static string Tab3Doc12End;
    public static string Tab3Doc12Count;

    public static string Tab3Doc13Start;
    public static string Tab3Doc13End;
    public static string Tab3Doc13Count;

    public static string Tab3Doc14Start;
    public static string Tab3Doc14End;
    public static string Tab3Doc14Count;
}