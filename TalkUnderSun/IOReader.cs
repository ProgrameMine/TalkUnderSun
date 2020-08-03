using System;
using System.IO;
using System.Text;

namespace TalkUnderSun
{
    public class IOReader
    {
        public void ReadFile()
        {
            //创建需要读取的数据的字节数组和字符数组
            byte[] byteData = new byte[20000];
            char[] charData = new char[20000];

            FileStream file = null;
            try
            {
                //打开一个当前 Program.cs 文件，此时读写文件的指针（或者说操作的光标）指向文件开头
                file = new FileStream(@"C:\Users\franz\Desktop\APIBankReconciliation.cs", FileMode.Open);
                //读写指针从开头往后移动10个字节
                file.Seek(3, SeekOrigin.Begin);
                //从当前读写指针的位置往后读取200个字节的数据到字节数组中
                file.Read(byteData, 0, 20000);
            }
            catch (Exception e)
            {
                Console.WriteLine("读取文件异常：{0}", e);
            }
            finally
            {
                //关闭文件流
                if (file != null) file.Close();
            }
            //创建一个编码转换器 解码器
            Decoder decoder = Encoding.UTF8.GetDecoder();
            //将字节数组转换为字符数组
            decoder.GetChars(byteData, 0, 20000, charData, 0);
            Console.WriteLine(charData);
        }

        public void WriteFile()
        {
            byte[] byteData;
            char[] charData;
            FileStream file = null;
            try
            {
                //在当前启动目录下的创建 aa.txt 文件
                file = new FileStream("aa.txt", FileMode.Create);
                //将“test write text to file”转换为字符数组并放入到 charData 中
                charData = "Test write text to file".ToCharArray();
                byteData = new byte[charData.Length];
                //创建一个编码器,将字符转换为字节
                Encoder encoder = Encoding.UTF8.GetEncoder();
                encoder.GetBytes(charData, 0, charData.Length, byteData, 0, true);
                file.Seek(0, SeekOrigin.Begin);
                //写入数据到文件中
                file.Write(byteData, 0, byteData.Length);

            }
            catch (Exception e)
            {
                Console.WriteLine("写入文件异常：{0}", e);
            }
            finally
            {
                if (file != null) file.Close();
            }
        }
    }
}
