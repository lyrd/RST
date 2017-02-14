using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RST
{
    class SavesIntoFile
    {
        public void SaveIntoFile(string content, string name, bool append)
        {
            using (StreamWriter str = new StreamWriter(name + ".txt", append))
            {
                str.WriteLine(content);
            }
        }

        public void SaveIntoFile(string content, string name, string fileExtension, bool append)
        {
            using (StreamWriter str = new StreamWriter(name + $".{fileExtension}", append))
            {
                str.WriteLine(content);
            }
        }

        public void SaveIntoFile<T>(T[] array, string name)
        {
            using (StreamWriter str = new StreamWriter(name + ".txt", true))
            {
                for (int i = 0; i < array.Length; i++)
                {
                    str.WriteLine(array[i] + "\t");
                }
            }
        }

        //public void SaveIntoFile(List<string> array, string name)
        //{
        //    using (StreamWriter str = new StreamWriter(name + ".txt", true))
        //    {
        //        for (int i = 0; i < array.Count; i++)
        //        {
        //            str.WriteLine(array[i] + "\t");
        //        }
        //    }
        //}
    }
}
